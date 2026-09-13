// D-Tools SI Cloud API — read side.
//
// Everything here reads from /Subscribe/*, which is a queue of records SI has published
// outward. Two things about that queue shape the code:
//
//   1. It is DELTA-based. A record appears only when it changes in SI, and each change
//      is a separate message. So the same project can appear many times and the newest
//      wins. Nothing is backfilled — a project untouched since publishing was enabled is
//      simply absent.
//
//   2. Reading with includeImported=true returns the whole queue rather than only
//      unseen messages. We do that deliberately and never call MarkAsImported: the queue
//      is our source of truth, and marking would hide records from any other integration
//      reading the same channel.
//
// Auth is a single API key header, generated per-user per-integration in
// SI Control Panel → Manage Integrations.

const BASE = 'https://api.d-tools.com/si';

export class DTools {
  constructor(apiKey) {
    if (!apiKey) throw new Error('D-Tools API key missing (DTOOLS_API_KEY).');
    this.key = apiKey;
  }

  async #get(path) {
    const res = await fetch(BASE + path, {
      headers: { 'X-DTSI-ApiKey': this.key, Accept: 'application/json' },
    });
    if (!res.ok) throw new Error(`GET ${path} -> ${res.status} ${(await res.text()).slice(0, 200)}`);
    return res.json();
  }

  /** Every project message in the queue, newest-per-project, approved and live. */
  async projects() {
    const all = [];
    for (let page = 1; ; page++) {
      const d = await this.#get(`/Subscribe/Projects?pageNumber=${page}&pageSize=200&includeImported=true`);
      const rows = d.Projects ?? [];
      all.push(...rows);
      if (all.length >= (d.TotalCount ?? 0) || rows.length === 0) break;
    }

    // Newest message per project id.
    const newest = new Map();
    for (const p of all) {
      const prev = newest.get(p.Id);
      if (!prev || (p.PublishedOnUnixTimeMs ?? 0) > (prev.PublishedOnUnixTimeMs ?? 0)) newest.set(p.Id, p);
    }

    return [...newest.values()].filter(p => p.Approved && !p.Archived && !p.Deleted);
  }

  /** Full project record. Large — includes every line item — so fetch sparingly. */
  projectDetail(id) { return this.#get(`/Subscribe/Projects/${encodeURIComponent(id)}`); }

  /** Scheduled tasks for one project. */
  tasks(projectId) { return this.#get(`/Subscribe/Tasks/ByProject/${encodeURIComponent(projectId)}`); }

  /** Every catalog message in the queue, newest first. Headers only -- no products. */
  async listCatalogs({ pageSize = 200 } = {}) {
    const all = [];
    for (let page = 1; ; page++) {
      const d = await this.#get(`/Subscribe/ProductCatalogs?pageNumber=${page}&pageSize=${pageSize}&includeImported=true`);
      const rows = d.Catalogs ?? [];
      all.push(...rows);
      if (all.length >= (d.TotalCount ?? 0) || rows.length === 0) break;
    }
    // Sorted here rather than trusting the API's order, because catalogVendors relies on
    // newest-wins when the same model appears in more than one message.
    all.sort((a, b) => (b.PublishedOnUnixTimeMs ?? 0) - (a.PublishedOnUnixTimeMs ?? 0));
    return { Catalogs: all };
  }

  /** One catalog message with its products. */
  getCatalog(id) { return this.#get(`/Subscribe/ProductCatalogs/${encodeURIComponent(id)}`); }
}

/* ---------- shaping ---------- */

// ART runs Trim and Finish as effectively the same visit -- a recent 430-item job carried
// 119.5 rough-in hours against 4.1 trim hours. Three segments would put a sliver on the
// wall nobody can read from across the room, so Trim folds into Finish for display.
// The catalog data is left alone; this is a presentation decision.
export function phaseOf(rawPhase) {
  const p = String(rawPhase || '').trim().toLowerCase();
  if (p.startsWith('rough')) return 'roughIn';
  if (p === 'trim' || p === 'finish') return 'finish';
  return null;
}

// Standard crew is two installers, eight hours -- so 16 labour hours is one day on site.
const HOURS_PER_DAY = 16;

/** Labour hours and estimated day-length per phase, from the project's line items. */
export function phaseEstimate(detail) {
  const hours = { roughIn: 0, finish: 0 };
  for (const it of detail.Items ?? []) {
    const k = phaseOf(it.Phase);
    if (k) hours[k] += Number(it.TotalLaborHours) || 0;
  }
  return {
    hours,
    days: {
      roughIn: round1(hours.roughIn / HOURS_PER_DAY),
      finish: round1(hours.finish / HOURS_PER_DAY),
    },
  };
}

// Tasks carry their own Progress, which ART uses to colour the SI calendar. The values
// mix three ideas -- phase (Prewire, Installing), work type (Security, Service,
// Lighting/Shades) and status (Not started). Only the phase ones are reliable enough to
// drive colour; the rest ride along as a label.
export function taskPhase(progress) {
  const p = String(progress || '').trim().toLowerCase();
  if (p.startsWith('prewire') || p.startsWith('rough')) return 'roughIn';
  if (p.startsWith('install') || p.startsWith('trim') || p.startsWith('finish')) return 'finish';
  return null;
}

/** Collapse a project's day-tasks into real date spans. */
export function taskSpans(tasks) {
  const out = [];
  for (const t of tasks ?? []) {
    const start = t.ScheduledStart || t.Start;
    const end = t.ScheduledEnd || t.End || start;
    if (!start) continue;
    out.push({
      id: t.Id,
      name: t.Name || '',
      start: dayOf(start),
      end: dayOf(end),
      progress: t.Progress || '',
      phase: taskPhase(t.Progress),
      pct: Number(t.PercentComplete) || 0,
      crew: (t.Resources ?? []).map(r => r.Name).filter(Boolean),
    });
  }
  return out.sort((a, b) => a.start.localeCompare(b.start));
}

// Date-only string. D-Tools returns local wall-clock times with no zone, so slicing is
// correct and safe -- parsing to a Date and back would shift days across timezones.
function dayOf(s) { return String(s).slice(0, 10); }
function round1(n) { return Math.round(n * 10) / 10; }

/* ---------- catalog vendor map ----------

   Why this exists: a project item snapshots the catalog when it is added, so its Vendor
   is frozen at that moment. Fixing a vendor in the catalog does NOT reach jobs already
   quoted -- you would otherwise have to edit it item by item in every project.

   Resolving by Model against the catalog instead means one catalog fix reaches every
   job, past and future. And because the catalog feed publishes products that changed,
   the products you fix are precisely the ones that arrive. The map fills in as the work
   gets done.

   The feed is deltas, so this is cumulative: merge what each run sees into what is
   already known rather than replacing it. */
export async function catalogVendors(si, { maxCatalogs = 100 } = {}) {
  const { Catalogs = [] } = await si.listCatalogs();
  const ordered = Catalogs.slice(0, maxCatalogs).reverse();   // newest N, applied oldest first so newest wins
  const map = new Map();
  const rqUse = new Map();                          // RQ SKU -> set of models using it
  for (const cat of ordered) {
    const full = await si.getCatalog(cat.Id);
    for (const p of full.Products ?? []) {
      if (!p.Model) continue;
      const key = p.Model.trim().toLowerCase();
      // CustomField1 holds the RepairQ SKU. It is the only bridge to the RQ stock
      // report, and it is heavily polluted -- one dead value, AVARNS000635, sits on 162
      // catalog products and does not exist in RQ at all. Track which SKUs are used by
      // more than one model so the untrustworthy ones can be ignored rather than
      // silently returning another product's stock.
      const rq = String(p.CustomField1 || '').trim();
      if (rq && rq !== '-') {
        const r = rq.toUpperCase();
        if (!rqUse.has(r)) rqUse.set(r, new Set());
        rqUse.get(r).add(key);
      }
      map.set(key, {
        model: p.Model,
        vendor: (p.Vendor && p.Vendor !== 'N/A') ? p.Vendor : '',
        cost: Number(p.UnitCost) || 0,
        rq: (rq && rq !== '-') ? rq : '',
        part: p.PartNumber || '',
      });
    }
  }
  // Second pass: blank out RQ SKUs that identify more than one product.
  for (const v of map.values()) {
    if (v.rq && (rqUse.get(v.rq.toUpperCase())?.size ?? 0) > 1) v.rq = '';
  }
  return map;
}

/* ---------- purchasing ---------- */

/* Which lines on a project are actually buyable.

   Three filters, all checked against real projects:

     TypeId 2        labour, not product. Appears as "Rough-In" / "Finish" / "Trim"
                     rows carrying fractional quantities and an hourly rate.
     IsOfe           the client supplies it.
     zero unit cost  the one that does the work. On a real 362-item job every single
                     zero-cost line was an "Existing ..." item -- flat panels, cable
                     boxes, mounts -- and all 351 costed lines survived.

   Deliberately NOT using the OFE flag to exclude "Existing ..." products: ticking it
   appears to strip cost from accessories, and the wire and plates that go with a
   client's own TV are things ART does supply and charge for. The cost filter achieves
   the same exclusion with no risk to a quote. */
export function orderLines(detail, vendorMap) {
  const byProduct = new Map();

  for (const it of detail.Items ?? []) {
    if (it.TypeId === 2) continue;                         // labour
    if (it.IsOfe || it.DoNotOrder) continue;
    if (!(Number(it.UnitCost) > 0)) continue;              // not purchased
    const qty = Number(it.Quantity) || 0;
    if (qty <= 0 || !it.Model) continue;

    const key = it.Model.trim().toLowerCase();
    // Catalog first, item second. The item's copy is a snapshot from whenever it was
    // added; the catalog is current. This is what lets one catalog fix reach every job.
    const cat = vendorMap?.get(key);
    const fromCatalog = cat?.vendor;
    const fromItem = (it.Vendor && it.Vendor !== 'N/A') ? it.Vendor : '';
    const row = byProduct.get(key) ?? {
      model: it.Model,
      manufacturer: it.Manufacturer || '',   // not "mfr" -- that key is the part number below
      desc: it.Description || '',
      vendor: fromCatalog || fromItem,
      vendorSource: fromCatalog ? 'catalog' : (fromItem ? 'project' : ''),
      part: it.PartNumber || '',
      cost: Number(it.UnitCost) || 0,
      phase: phaseOf(it.Phase) || 'finish',   // unassigned rides with the later phase
      qty: 0,
      // Bulk wire is measured, not counted. In the API payload Quantity on a bulk item
      // is already the FOOTAGE (a 16-2 run comes through as Quantity 125), so the sum is
      // correct -- but "3245" means feet, not pieces, and ordering 33 boxes because a
      // job has 33 runs would be badly wrong. The flag drives the unit on screen.
      // (Note the CSV export models this differently: one row per run with a separate
      // Wire Length column. This code reads the API.)
      bulk: !!(it.BulkWire || it.BulkItem),
      ordered: !!(it.OrderedDateTicks > 0),
      // Keys the Orders page uses to look this product up in a pasted RQ stock report.
      // rq is blank when the catalog's value is shared across products and so cannot
      // identify stock; mfr is the fallback, matched against the report's
      // Manufacturer SKU column.
      rq: cat?.rq || '',
      mfr: (cat?.part || it.PartNumber || ''),
    };
    row.qty += qty;
    byProduct.set(key, row);
  }

  // Round up: you cannot order 3.4 connectors.
  return [...byProduct.values()].map(r => ({ ...r, qty: Math.ceil(r.qty) }));
}
