// QuickBooks Online client.
//
// STATUS: not live yet. QBO Plus is planned for purchase ~October 2026, ahead of the
// Dec 31 2026 handover. Until then `stubInventory()` supplies fake numbers so the
// D-Tools half of the sync can be built and tested end to end.
//
// ---------------------------------------------------------------------------
// Mapping the three SI custom fields onto QuickBooks
// ---------------------------------------------------------------------------
// Only one of the three is a native QuickBooks field. Worth knowing before wiring
// this up, because the other two cost extra API calls:
//
//   Stock     -> Item.QtyOnHand. Native, one query, trivially available.
//
//   Committed -> No native equivalent. QBO Plus has no Sales Orders (that is a
//                Desktop feature). The closest analogue is quantity sitting on open
//                Estimates that have not yet been converted to an Invoice, which means
//                querying Estimate objects and summing their line items per SKU.
//
//   Ordered   -> No native equivalent. Derived by querying PurchaseOrder objects still
//                in an open//pending state and summing line items per SKU.
//
// So Stock works on day one; Committed and Ordered need aggregation passes that are
// worth validating against real data before trusting them on a salesperson's screen.
//
// ---------------------------------------------------------------------------
// The refresh-token trap
// ---------------------------------------------------------------------------
// Intuit refresh tokens ROTATE: every refresh returns a new refresh token and
// invalidates the old one. They also expire after ~100 days of disuse. A scheduled
// job that does not persist the new token somewhere durable will work perfectly in
// testing and then die silently weeks later. `onTokenRefresh` exists for exactly this
// -- see sync.mjs, which persists to a SharePoint list via Graph.
//
// This is the same class of bug as the `offline_access` scope fix on the Pull Log.

const TOKEN_URL = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer';
const API_BASE = {
  production: 'https://quickbooks.api.intuit.com',
  sandbox: 'https://sandbox-quickbooks.api.intuit.com',
};

export class QuickBooks {
  /**
   * @param {object}   opts
   * @param {string}   opts.clientId
   * @param {string}   opts.clientSecret
   * @param {string}   opts.refreshToken   Current refresh token (rotates on every use).
   * @param {string}   opts.realmId        QBO company id.
   * @param {'production'|'sandbox'} [opts.environment]
   * @param {(token: string) => Promise<void>} [opts.onTokenRefresh]
   *        Called with the NEW refresh token after each refresh. Must persist it.
   */
  constructor({ clientId, clientSecret, refreshToken, realmId, environment = 'production', onTokenRefresh }) {
    Object.assign(this, { clientId, clientSecret, refreshToken, realmId, environment, onTokenRefresh });
    this.accessToken = null;
  }

  async #refreshAccessToken() {
    const basic = Buffer.from(`${this.clientId}:${this.clientSecret}`).toString('base64');
    const res = await fetch(TOKEN_URL, {
      method: 'POST',
      headers: {
        'Authorization': `Basic ${basic}`,
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json',
      },
      body: new URLSearchParams({
        grant_type: 'refresh_token',
        refresh_token: this.refreshToken,
      }),
    });

    if (!res.ok) {
      throw new Error(`QBO token refresh failed: HTTP ${res.status} ${await res.text()}`);
    }

    const data = await res.json();
    this.accessToken = data.access_token;

    // Rotation: persist immediately, before any other call can fail and lose it.
    if (data.refresh_token && data.refresh_token !== this.refreshToken) {
      this.refreshToken = data.refresh_token;
      await this.onTokenRefresh?.(data.refresh_token);
    }
    return this.accessToken;
  }

  async #query(sql) {
    if (!this.accessToken) await this.#refreshAccessToken();

    const url = `${API_BASE[this.environment]}/v3/company/${this.realmId}/query?query=${encodeURIComponent(sql)}`;
    const res = await fetch(url, {
      headers: { 'Authorization': `Bearer ${this.accessToken}`, 'Accept': 'application/json' },
    });

    if (!res.ok) throw new Error(`QBO query failed: HTTP ${res.status} ${await res.text()}`);
    return (await res.json()).QueryResponse ?? {};
  }

  // QBO caps query results at 1000 rows, so page with STARTPOSITION/MAXRESULTS.
  async #queryAll(select, entity) {
    const out = [];
    const page = 1000;
    for (let start = 1; ; start += page) {
      const qr = await this.#query(`${select} STARTPOSITION ${start} MAXRESULTS ${page}`);
      const rows = qr[entity] ?? [];
      out.push(...rows);
      if (rows.length < page) break;
    }
    return out;
  }

  /**
   * Returns Map<sku, { stock, committed, ordered }> keyed by the item's SKU.
   *
   * The SKU must equal the D-Tools SI Model string exactly (dashes and spaces
   * included, e.g. "AN-520-AP-I-1"). Items with no SKU are skipped -- they cannot be
   * matched, and silently guessing from Name would produce wrong stock numbers.
   */
  async getInventory() {
    const items = await this.#queryAll(
      "SELECT Id, Name, Sku, QtyOnHand FROM Item WHERE Type = 'Inventory' AND Active = true",
      'Item',
    );

    const inv = new Map();
    for (const it of items) {
      if (!it.Sku) continue;
      inv.set(String(it.Sku).trim().toLowerCase(), {
        stock: Number(it.QtyOnHand ?? 0),
        committed: 0,
        ordered: 0,
      });
    }

    await this.#addOrdered(inv);
    await this.#addCommitted(inv);
    return inv;
  }

  // Quantity on open purchase orders, summed per SKU.
  async #addOrdered(inv) {
    const pos = await this.#queryAll(
      "SELECT Id, POStatus, Line FROM PurchaseOrder WHERE POStatus = 'Open'",
      'PurchaseOrder',
    );
    for (const po of pos) accumulate(inv, po.Line, 'ordered');
  }

  // Quantity on open (not yet invoiced) estimates, summed per SKU.
  // QBO Plus has no Sales Order object; estimates are the nearest thing to a
  // commitment. Validate this against how the team actually quotes before trusting it.
  async #addCommitted(inv) {
    const ests = await this.#queryAll(
      "SELECT Id, TxnStatus, Line FROM Estimate WHERE TxnStatus IN ('Pending', 'Accepted')",
      'Estimate',
    );
    for (const e of ests) accumulate(inv, e.Line, 'committed');
  }
}

function accumulate(inv, lines, field) {
  for (const line of lines ?? []) {
    const detail = line.SalesItemLineDetail ?? line.ItemBasedExpenseLineDetail;
    const sku = detail?.ItemRef?.name;
    if (!sku) continue;
    const row = inv.get(String(sku).trim().toLowerCase());
    if (row) row[field] += Number(detail.Qty ?? 0);
  }
}

/**
 * Fake inventory for testing the D-Tools half before QuickBooks exists.
 * Given SI models, returns deterministic quantities so repeated runs are stable
 * and a genuine change is distinguishable from random noise.
 */
export function stubInventory(models) {
  const inv = new Map();
  for (const m of models) {
    const key = String(m).trim().toLowerCase();
    let h = 0;
    for (const ch of key) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
    inv.set(key, {
      stock: h % 40,
      committed: h % 7,
      ordered: h % 5,
    });
  }
  return inv;
}
