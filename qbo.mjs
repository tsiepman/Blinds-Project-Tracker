// QuickBooks Online — the accounting side of the fulfilment loop.
//
// Why this lives in the sync and not in the Pull Log: Intuit's token exchange needs the
// client secret as HTTP Basic auth. Anything in a page served from GitHub Pages is
// readable by anyone who views source, so the browser can never hold it. The installers'
// tools stay SharePoint-only; everything that touches QuickBooks happens here, where the
// secrets already live.
//
// The refresh token ROTATES on every use. Intuit hands back a new one with each refresh
// and the old one stops working, so it has to be written somewhere durable before the
// next run. A GitHub Action cannot write its own secrets, so the token is kept in
// SharePoint and passed in through the store below. Get this wrong and the integration
// works for weeks and then dies quietly, which is the worst way for it to fail.
//
// Intuit changed the refresh-token validity policy in November 2025 and the old "100 days"
// figure is no longer something to rely on. The design does not depend on a lifetime: it
// persists every rotation, and when Intuit finally refuses with invalid_grant it says so
// loudly, because the fix is a human re-authorising and nothing else will do.

const TOKEN_URL = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer';
const API = {
  sandbox: 'https://sandbox-quickbooks.api.intuit.com',
  production: 'https://quickbooks.api.intuit.com',
};
// Minor versions below 75 were retired on 1 August 2025. Pinned rather than floating, so
// Intuit changing the default cannot change our behaviour without us choosing it.
const MINOR_VERSION = '75';

export class ReauthorizeNeeded extends Error {
  constructor(msg) { super(msg); this.name = 'ReauthorizeNeeded'; }
}

export class Qbo {
  /* store: { read(): Promise<{refreshToken, realmId}>, write(obj): Promise<void> }
     Injected so this module knows nothing about SharePoint. */
  constructor({ clientId, clientSecret, environment = 'production', store }) {
    for (const [k, v] of Object.entries({ clientId, clientSecret, store })) {
      if (!v) throw new Error(`Qbo: ${k} is missing.`);
    }
    Object.assign(this, { clientId, clientSecret, store });
    this.base = API[environment] || API.production;
    this.environment = environment;
    this.access = null;
    this.accessExpires = 0;
    this.realmId = null;
  }

  async #refresh() {
    const saved = await this.store.read();
    if (!saved?.refreshToken) {
      throw new ReauthorizeNeeded('No QuickBooks refresh token stored. Authorise once and save the token.');
    }
    this.realmId = saved.realmId || this.realmId;

    const res = await fetch(TOKEN_URL, {
      method: 'POST',
      headers: {
        Authorization: 'Basic ' + Buffer.from(`${this.clientId}:${this.clientSecret}`).toString('base64'),
        'Content-Type': 'application/x-www-form-urlencoded',
        Accept: 'application/json',
      },
      body: new URLSearchParams({ grant_type: 'refresh_token', refresh_token: saved.refreshToken }),
    });
    const text = await res.text();
    if (!res.ok) {
      // invalid_grant means the stored token is spent or revoked. No amount of retrying
      // fixes it -- somebody has to click through Intuit's consent screen again.
      if (/invalid_grant/i.test(text)) {
        throw new ReauthorizeNeeded(
          'QuickBooks refused the stored refresh token (invalid_grant). Re-authorise and save a new token — ' +
          'nothing will post until that is done.');
      }
      throw new Error(`QuickBooks token refresh failed: ${res.status} ${text.slice(0, 200)}`);
    }
    const d = JSON.parse(text);
    this.access = d.access_token;
    // A minute of slack so a long run cannot straddle the expiry.
    this.accessExpires = Date.now() + ((d.expires_in ?? 3600) - 60) * 1000;

    // Persist FIRST, use second. If the write fails we would rather fail now than carry on
    // with a token the store has never seen -- that is how the chain gets broken.
    if (d.refresh_token && d.refresh_token !== saved.refreshToken) {
      await this.store.write({ ...saved, refreshToken: d.refresh_token, rotatedAt: new Date().toISOString() });
    }
    return this.access;
  }

  async #token() {
    if (this.access && Date.now() < this.accessExpires) return this.access;
    return this.#refresh();
  }

  async req(method, path, body, attempt = 1) {
    const token = await this.#token();
    const realm = this.realmId || (await this.store.read())?.realmId;
    if (!realm) throw new Error('Qbo: no realmId (company id) stored.');
    const join = path.includes('?') ? '&' : '?';
    const res = await fetch(`${this.base}/v3/company/${realm}${path}${join}minorversion=${MINOR_VERSION}`, {
      method,
      headers: {
        Authorization: 'Bearer ' + token,
        Accept: 'application/json',
        ...(body ? { 'Content-Type': 'application/json' } : {}),
      },
      ...(body ? { body: JSON.stringify(body) } : {}),
    });

    // 401 once means the access token died early; refresh and try again. Twice means
    // something else is wrong and retrying would just hammer Intuit.
    if (res.status === 401 && attempt === 1) {
      this.access = null;
      return this.req(method, path, body, 2);
    }
    // Intuit throttles per realm. Back off rather than dropping the pull on the floor --
    // an unposted pull is product out the door that nobody invoiced.
    if ((res.status === 429 || res.status >= 500) && attempt < 4) {
      await new Promise(r => setTimeout(r, attempt * 2000));
      return this.req(method, path, body, attempt + 1);
    }
    const text = await res.text();
    if (!res.ok) throw new Error(`QuickBooks ${method} ${path} -> ${res.status} ${text.slice(0, 300)}`);
    return text ? JSON.parse(text) : null;
  }

  /* The query endpoint takes a SQL-ish string. Single quotes inside values have to be
     doubled or the query breaks on any customer with an apostrophe in their name, which
     in this book is a matter of time. */
  async query(sql) {
    const d = await this.req('GET', '/query?query=' + encodeURIComponent(sql));
    const qr = d?.QueryResponse ?? {};
    const key = Object.keys(qr).find(k => Array.isArray(qr[k]));
    return key ? qr[key] : [];
  }

  /* Reads everything, not just the first page. QBO caps a query at 1000 rows and simply
     stops -- no flag, no warning -- so a shop with more estimates than that would silently
     see only some of them. */
  async queryAll(sql, pageSize = 500) {
    const out = [];
    for (let start = 1; ; start += pageSize) {
      const page = await this.query(`${sql} STARTPOSITION ${start} MAXRESULTS ${pageSize}`);
      out.push(...page);
      if (page.length < pageSize) break;
    }
    return out;
  }

  estimates({ since } = {}) {
    const where = since ? ` WHERE MetaData.LastUpdatedTime > '${since}'` : '';
    return this.queryAll(`SELECT * FROM Estimate${where}`);
  }
  invoice(id) { return this.query(`SELECT * FROM Invoice WHERE Id = '${String(id).replace(/'/g, "''")}'`).then(r => r[0] || null); }
  items() { return this.queryAll('SELECT * FROM Item'); }
  customers() { return this.queryAll('SELECT * FROM Customer'); }

  createInvoice(invoice) { return this.req('POST', '/invoice', invoice).then(r => r?.Invoice); }

  /* Full-object update, not a sparse one. Sparse updates are documented but behave
     inconsistently across entity types, and an invoice whose lines are half-replaced is a
     bill that goes out wrong. Read, change, send the whole thing back with its SyncToken. */
  updateInvoice(invoice) { return this.req('POST', '/invoice', invoice).then(r => r?.Invoice); }

  /** Proves the credentials and the company before anything is written. */
  async whoami() {
    const info = await this.req('GET', '/companyinfo/' + (this.realmId || (await this.store.read())?.realmId));
    const c = info?.CompanyInfo ?? {};
    return { name: c.CompanyName || '', legalName: c.LegalName || '', country: c.Country || '', environment: this.environment };
  }
}

/* Token storage in SharePoint.
   A refresh token is full access to the accounting system, so the list it lives in should
   have its permission inheritance broken -- app plus Trevor, nobody else. The sync reaches
   it with app-only credentials either way; the restriction is about people. */
export function sharePointTokenStore(graph, listName = 'SyncConfig', title = 'qbo') {
  let cached = null;
  return {
    async read() {
      if (cached) return cached;
      const rows = await graph.items(listName).catch(() => []);
      const row = rows.find(i => (i.fields ?? {}).Title === title);
      if (!row) return null;
      let data = {};
      try { data = JSON.parse(row.fields.DataJson || '{}'); } catch { data = {}; }
      cached = { ...data, _itemId: row.id };
      return cached;
    },
    async write(obj) {
      const { _itemId, ...data } = obj;
      const fields = { Title: title, DataJson: JSON.stringify(data) };
      if (_itemId) await graph.update(listName, _itemId, fields);
      else {
        const made = await graph.create(listName, fields);
        obj._itemId = made?.id;
      }
      cached = { ...data, _itemId: _itemId || obj._itemId };
    },
  };
}
