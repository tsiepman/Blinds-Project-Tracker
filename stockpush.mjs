// RepairQ stock -> the D-Tools catalog, so a salesman speccing a job sees what is held.
//
//   Brian loads the RQ report on the Orders page
//     -> the page writes ARTStock/rq-stock.json to the ART43 document library
//     -> this reads it, matches each product to the catalog, and publishes
//        CustomField6 (Stock), CustomField7 (Committed), CustomField8 (On order).
//
// RepairQ stays the master. This is a once-a-day copy for people quoting work, not a
// live availability check, and nothing here reads back into the order figures.
//
// Two things make this careful rather than simple:
//
//   1. A catalog publish writes the WHOLE product record. Each product is therefore sent
//      back exactly as SI published it, with only those three fields changed. Sending a
//      stripped-down record would blank cost, vendor and labour across the catalog.
//   2. It only sends products whose numbers actually moved, so an unchanged shelf
//      publishes nothing.
//
// SI Control Panel -> Manage Integrations must have catalog updates enabled. With it off
// the call is accepted and silently does nothing, which looks exactly like success.

const STOCK_FILE = 'ARTStock/rq-stock.json';
const STOCK_LIST = 'StockSnapshot';

// Which custom fields hold what, as set up in SI. 6, 7 and 8 are the numeric ones.
const F_STOCK = 'CustomField6';
const F_COMMITTED = 'CustomField7';
const F_ONORDER = 'CustomField8';

// A whole catalog in one POST would be megabytes; SI takes them in messages anyway.
const CHUNK = 250;

export async function pushStock(si, graph, getVendorMap, { dryRun = false } = {}) {
  let raw;
  try { raw = await graph.fileText(STOCK_FILE); }
  catch (e) { console.log(`\nstock push skipped: ${e.message.slice(0, 160)}`); return; }
  if (!raw) return;                                  // no report has been loaded yet

  let file;
  try { file = JSON.parse(raw); } catch { console.log('\nstock push skipped: the stock file is not readable'); return; }
  const rows = file.rows ?? [];
  if (!rows.length) return;

  // Already done for this report? The row carries the report's own timestamp.
  const items = await graph.items(STOCK_LIST).catch(() => []);
  const pushRow = items.find(i => (i.fields ?? {}).Title === 'pushed');
  if (pushRow && String(pushRow.fields?.AsOf || '') === String(file.asOf || '') && !process.env.STOCK_PUSH_FORCE) {
    return;                                          // same report as last time
  }

  // Catalog, with each product's full record. Matching is by RepairQ SKU (CustomField1),
  // then part number -- the same two keys the Orders page uses, so what a salesman sees in
  // SI and what the Orders page nets against can never disagree about which product it is.
  const cat = await getVendorMap();
  const byRq = new Map(), byPart = new Map();
  for (const v of cat.values()) {
    if (v.rq) byRq.set(norm(v.rq), v);               // already blanked when shared
    if (v.part) {
      const k = norm(v.part);
      // A part number on two products identifies neither, same rule as the RQ SKU.
      byPart.set(k, byPart.has(k) ? null : v);
    }
  }

  const send = [];
  let matched = 0, unmatched = 0, unchanged = 0;
  for (const r of rows) {
    const v = byRq.get(norm(r.sku)) || byPart.get(norm(r.mfr)) || null;
    if (!v || !v.full) { unmatched++; continue; }
    matched++;
    const p = v.full;
    const want = { [F_STOCK]: num(r.stock), [F_COMMITTED]: num(r.committed), [F_ONORDER]: num(r.onOrder) };
    // Committed is only written when the report carries the column; a missing column must
    // not overwrite a real figure in SI with a zero.
    if (r.committed == null) delete want[F_COMMITTED];
    if (Object.entries(want).every(([k, n]) => num(p[k]) === n)) { unchanged++; continue; }
    send.push({ ...p, ...want });
  }

  console.log(`\nstock push: ${rows.length} products in the report  ${matched} matched to the catalog` +
    `  ${unmatched} not in the catalog  ${unchanged} already right  ${send.length} to update`);
  if (!send.length) { await stamp(); return; }

  // A limit for the first real run: set STOCK_PUSH_LIMIT=1, check that product in SI, then
  // take the limit off.
  const limit = Number(process.env.STOCK_PUSH_LIMIT) || 0;
  const batch = limit ? send.slice(0, limit) : send;
  if (limit) console.log(`  (STOCK_PUSH_LIMIT=${limit} -- only the first ${batch.length} will be sent)`);

  for (const p of batch.slice(0, 5)) {
    console.log(`    ${p.Model}: stock ${p[F_STOCK]}${p[F_COMMITTED] === undefined ? '' : `  committed ${p[F_COMMITTED]}`}  on order ${p[F_ONORDER]}`);
  }
  if (batch.length > 5) console.log(`    … and ${batch.length - 5} more`);

  if (dryRun) { console.log('  (dry run -- nothing sent to D-Tools)'); return; }

  let sent = 0;
  for (let i = 0; i < batch.length; i += CHUNK) {
    const part = batch.slice(i, i + CHUNK);
    try {
      await si.publishCatalog({
        Name: file.catalogName || 'RepairQ Stock',
        Description: `Stock, committed and on order from RepairQ, as of ${String(file.asOf || '').slice(0, 10)}`,
        IsMetric: false,
        TotalProductsCount: part.length,
        Products: part,
      });
      sent += part.length;
    } catch (e) {
      console.log(`  ! stock push failed after ${sent}: ${e.message.slice(0, 200)}`);
      return;                                        // leave the stamp alone so it retries
    }
  }
  console.log(`  ${sent} products published to the catalog`);
  if (!limit) await stamp();                         // a limited run must not count as done

  async function stamp() {
    if (dryRun) return;
    const fields = { Title: 'pushed', AsOf: String(file.asOf || ''), DataJson: JSON.stringify({ at: new Date().toISOString(), sent: send.length }) };
    try {
      if (pushRow) await graph.update(STOCK_LIST, pushRow.id, fields);
      else await graph.create(STOCK_LIST, fields);
    } catch (e) { console.log(`  ! push marker not saved: ${e.message.slice(0, 160)}`); }
  }
}

function norm(s) { return String(s ?? '').replace(/[^A-Za-z0-9]/g, '').toUpperCase(); }
function num(n) { const x = Number(n); return Number.isFinite(x) ? x : 0; }
