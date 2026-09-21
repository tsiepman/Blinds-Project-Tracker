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

  /** Every catalog message in the queue, OLDEST first. Headers only -- no products. */
  async listCatalogs({ pageSize = 200 } = {}) {
    const all = [];
    for (let page = 1; ; page++) {
      const d = await this.#get(`/Subscribe/ProductCatalogs?pageNumber=${page}&pageSize=${pageSize}&includeImported=true`);
      const rows = d.Catalogs ?? [];
      all.push(...rows);
      if (all.length >= (d.TotalCount ?? 0) || rows.length === 0) break;
    }
    // Catalog headers carry only PublishedOn ("2026-09-12T23:26:00") -- not the
    // PublishedOnUnixTimeMs that project messages have, so sorting on that was a no-op
    // that only worked because the API happens to return newest first. ISO strings sort
    // correctly as text, which avoids any timezone parsing.
    all.sort((a, b) => String(a.PublishedOn || '').localeCompare(String(b.PublishedOn || '')));
    return { Catalogs: all };
  }

  /** One catalog message with its products. */
  getCatalog(id) { return this.#get(`/Subscribe/ProductCatalogs/${encodeURIComponent(id)}`); }

  /** Service order messages in the queue (service calls, and the crew's day rows). */
  async serviceOrders({ pageSize = 200 } = {}) {
    const all = [];
    for (let page = 1; ; page++) {
      const d = await this.#get(`/Subscribe/ServiceOrders?pageNumber=${page}&pageSize=${pageSize}&includeImported=true`);
      const rows = d.ServiceOrders ?? [];
      all.push(...rows);
      if (all.length >= (d.TotalCount ?? 0) || rows.length === 0) break;
    }
    const newest = new Map();
    for (const s of all) newest.set(s.Id, s);
    return [...newest.values()];
  }

  /** One service order with its crew, dates and notes. */
  getServiceOrder(id) { return this.#get(`/Subscribe/ServiceOrders/${encodeURIComponent(id)}`); }
}

/* ---------- field visits ----------

   A record of who was scheduled where, written down as it happens. The API only ever
   shows the CURRENT state: reschedule a task and yesterday's version is gone, so there
   would be nothing left to check a work-order write-up against. This captures each
   visit while it is still there.

   Both sources look the same once shaped: a task on a project, or a service order.   */

function day(s) { return s ? String(s).slice(0, 10) : ''; }
function clock(s) { return s ? String(s).slice(11, 16) : ''; }
function crewOf(x) { return (x.Resources ?? []).map(r => r.Name).filter(Boolean); }
function addressOf(a) {
  if (!a) return '';
  return [a.Street1, a.City].filter(Boolean).join(', ');
}

export function visitFromTask(t, project) {
  if (!t.ScheduledStart) return null;
  return {
    visitId: t.Id,
    kind: 'task',
    number: String(t.Number ?? ''),
    jobNumber: project?.Number || t.ProjectNumber || '',
    jobName: project?.Name || t.Project || '',
    client: t.Client || project?.Client || '',
    site: addressOf(t.SiteAddress),
    name: t.Name || '',
    date: day(t.ScheduledStart),
    from: clock(t.ScheduledStart),
    to: clock(t.ScheduledEnd),
    crew: crewOf(t),
    progress: t.Progress || '',
    pct: Number(t.PercentComplete) || 0,
    instructions: t.Description || '',
    updatedOn: t.UpdatedOn || '',
    updatedBy: t.UpdatedBy || '',
  };
}

export function visitFromServiceOrder(s) {
  if (!s.ScheduledStart) return null;
  return {
    visitId: s.Id,
    kind: 'service',
    number: String(s.Number ?? ''),
    jobNumber: s.ProjectNumber || '',
    jobName: s.Project || '',
    client: s.Client || '',
    site: addressOf(s.SiteAddress),
    name: s.Name || '',
    date: day(s.ScheduledStart),
    from: clock(s.ScheduledStart),
    to: clock(s.ScheduledEnd),
    crew: crewOf(s),
    progress: s.Progress || '',
    pct: Number(s.PercentComplete) || 0,
    // Service orders carry both: Description is why it was raised, Notes is what to do.
    instructions: [s.Description, s.Notes].filter(Boolean).join('\n'),
    updatedOn: s.UpdatedOn || '',
    updatedBy: s.UpdatedBy || '',
  };
}

/** Keep visits scheduled from `back` days ago to `ahead` days out. */
export function visitInWindow(v, back = 30, ahead = 14) {
  if (!v?.date) return false;
  const t = new Date(); t.setHours(0, 0, 0, 0);
  const d = new Date(v.date + 'T00:00:00');
  const diff = Math.round((d - t) / 86400000);
  return diff >= -back && diff <= ahead;
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

// Standard crew is two installers, but nobody gets a full eight hours on the job once
// travel, pickups and setup come out -- so 13 labour hours is one day on site.
// The board and Orders pages use the same figure (HOURS_PER_DAY in each).
const HOURS_PER_DAY = 13;

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
// Lighting/Shades) and status (Not started). Rough-in is anything marked Prewire, or
// named for rough-in ("Rough in Placeholder"); everything else is finish work, so a
// Lighting/Shades or Security day still shows and still drives ordering.
// The board and Orders pages repeat this rule so older saved tasks read the same way.
export function taskPhase(progress, name) {
  const p = String(progress || '').trim().toLowerCase();
  const n = String(name || '').toLowerCase();
  if (p.startsWith('prewire') || p.startsWith('rough') || /rough|pre-?wire/.test(n)) return 'roughIn';
  return 'finish';
}

// Salespeople add "Rough in Placeholder" / "Finish Placeholder" tasks when a quote is
// approved. They mark roughly when a phase starts; the length comes from quoted hours.
// Once a phase has any real task, its placeholder is ignored.
export function isPlaceholder(name) { return /placeholder|tentative/i.test(String(name || '')); }

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
      phase: taskPhase(t.Progress, t.Name),
      placeholder: isPlaceholder(t.Name),
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

   The feed is deltas, so EVERY message has to be read, oldest first, newest copy of a
   product winning. Do not cap it to the newest N: the base catalog is the OLDEST data in
   the queue -- the first publish on 2026-04-02 carried 5,217 products and a republish on
   04-15 another 3,727, while daily messages since carry 5 to 30. A cap of 100 was
   silently dropping both once the queue passed 100 messages, leaving only products
   edited since late April. */
export async function catalogVendors(si, catalogs) {
  const ordered = catalogs ?? (await si.listCatalogs()).Catalogs;   // oldest first
  const map = new Map();
  const byId = new Map();                           // product id -> the key it lives under
  for (const cat of ordered) {
    const full = await si.getCatalog(cat.Id);
    // Deletions come as ids on the message that removed them. Ignoring them leaves
    // products that no longer exist in D-Tools sitting in the map for ever -- which is
    // how a discontinued product went on holding a RepairQ SKU and got the live product
    // sharing it blanked, showing "?" instead of its stock.
    for (const id of full.DeletedProductIds ?? []) {
      const key = byId.get(id);
      if (key && map.get(key)?.id === id) map.delete(key);
      byId.delete(id);
    }
    for (const p of full.Products ?? []) {
      if (!p.Model) continue;
      const key = p.Model.trim().toLowerCase();
      // A product that was renamed leaves its old model key behind, so drop that too.
      const wasAt = p.Id ? byId.get(p.Id) : null;
      if (wasAt && wasAt !== key) map.delete(wasAt);
      if (p.Id) byId.set(p.Id, key);
      // CustomField1 holds the RepairQ SKU. It is the only bridge to the RQ stock
      // report, and it is heavily polluted -- one dead value, AVARNS000635, sits on 162
      // catalog products and does not exist in RQ at all.
      const rq = String(p.CustomField1 || '').trim();
      map.set(key, {
        id: p.Id || '',
        model: p.Model,
        vendor: (p.Vendor && p.Vendor !== 'N/A') ? p.Vendor : '',
        cost: Number(p.UnitCost) || 0,
        rq: (rq && rq !== '-') ? rq : '',
        part: p.PartNumber || '',
        // Read from the catalog for the same reason vendor is: ticking Do Not Order on
        // ARTPART or PERDIEM once should clear it from every job, not just new ones.
        doNotOrder: !!p.DoNotOrder,
      });
    }
  }
  // Blank out RQ SKUs that identify more than one product, so an untrustworthy one is
  // ignored rather than silently returning another product's stock. Counted on the final
  // map, not while reading -- a SKU someone has since corrected must not still count
  // against the product it used to be on.
  const rqUse = new Map();                          // RQ SKU -> the models using it
  for (const v of map.values()) {
    if (!v.rq) continue;
    const k = v.rq.toUpperCase();
    if (!rqUse.has(k)) rqUse.set(k, []);
    rqUse.get(k).push(v.model);
  }
  for (const v of map.values()) {
    if (v.rq && rqUse.get(v.rq.toUpperCase()).length > 1) v.rq = '';
  }
  // Named in the sync log: each one is a catalog fix that gives a product its stock back.
  map.sharedRq = [...rqUse.entries()].filter(([, ms]) => ms.length > 1)
    .sort((a, b) => b[1].length - a[1].length)
    .map(([sku, ms]) => ({ sku, models: ms }));
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
    if (cat?.doNotOrder) continue;                         // flagged in the catalog since quoting
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
