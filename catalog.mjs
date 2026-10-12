// Catalog health and the QuickBooks item list.
//
//   Run the workflow with "Catalog report" ticked. Nothing here runs on a normal tick --
//   it reads all ~5,400 catalog products and that is not work for a twenty-minute cron.
//
// Four files land in catalog-report/ and come back as a workflow artifact:
//
//   catalog-issues.csv    the cleanup worklist, worst first
//   qbo-items-N.csv       QuickBooks item import, 1000 rows a file, bad rows left out
//   qbo-suppliers.csv     the supplier names, which have to exist in QuickBooks FIRST
//   catalog-master.csv    every product with every field, for reference
//
// Why this exists: the Model is the join key across three systems. A salesman types it in
// D-Tools, it becomes the SKU in QuickBooks, and an installer scans it in the Pull Log. If
// one product can be two Models, or two products one SKU, a scan resolves to the wrong
// thing and the invoice is wrong. Cleaning that up in the catalog is cheap; cleaning it up
// in QuickBooks after transactions exist is not.
//
// Nothing here writes to D-Tools, SharePoint or QuickBooks. It only reads and reports.

import { writeFileSync, mkdirSync } from 'node:fs';

const OUT = 'catalog-report';
const STOCK_FILE = 'ARTStock/rq-stock.json';
// QuickBooks rejects a name or SKU over 100 characters outright.
const QBO_MAX = 100;
// QuickBooks caps an item import at 1000 rows a file.
const QBO_ROWS = 1000;

/* The QuickBooks import header row, taken verbatim from the sample file this account
   serves at Gear > Import data > Products and Services > Download a sample file. Not from
   memory: QuickBooks varies these by version, and a header it does not recognise is left
   unmapped rather than refused, so the column goes in silently empty.

   Account columns resolve to '' on purpose. An inventory item needs an income, an expense
   and an inventory asset account, and those come from Kevin's chart of accounts -- filling
   them before that is settled means re-importing every row. The headers are here so
   December is a fill-in rather than a rebuild. */
const QBO_COLUMNS = [
  ['Product/Service name', p => p.model],
  ['Category', p => p.category],
  ['Item type', p => p.type],
  ['SKU', p => p.model],
  ['Sales description', p => p.desc || p.model],
  ['Sales price/rate', p => p.price || ''],
  ['Income account', () => ''],
  ['Purchase description', p => p.desc || p.model],
  ['Purchase cost', p => p.cost || ''],
  ['Expense account', () => ''],
  // Only meaningful on Inventory rows, and only when a date is given -- see QTY_AS_OF.
  ['Quantity on hand', p => (p.type === 'Inventory' && QTY_AS_OF) ? 0 : ''],
  ['Quantity as of date', p => (p.type === 'Inventory' && QTY_AS_OF) ? QTY_AS_OF : ''],
  ['Reorder point', () => ''],
  ['Preferred Vendor', p => p.vendor],
  ['Inventory asset account', () => ''],
];

/* QuickBooks wants a quantity and an as-of date on an Inventory row. Both are left blank
   unless CATALOG_QTY_AS_OF is set (MM/DD/YYYY), because the real opening count comes from
   the stocktake at closing, not from RepairQ -- it is entered afterwards as an inventory
   adjustment. Set it to 01/01/2027 for the live import so the rows load at zero and the
   count sets them; leave it unset for a dry look at the data. */
const QTY_AS_OF = String(process.env.CATALOG_QTY_AS_OF || '').trim();

/* usedModels: every model that appears as an order line on a project the sync has seen.
   It answers the question the raw count cannot -- of 5,400 catalog products, which ones has
   anybody actually quoted? A product nobody has used in the synced history is a row in the
   QuickBooks item list for ever and a product to clean up now, for no return. Reported, not
   filtered: where to draw that line is a decision, not a rule. */
export async function catalogReport(graph, getVendorMap, usedModels = new Set()) {
  const cat = await getVendorMap();

  /* RepairQ, for two separate jobs.
     1. Which products RepairQ actually tracks -- that is what decides inventory vs
        non-inventory in QuickBooks, rather than someone guessing product by product.
     2. Which RepairQ SKUs in the catalog are dead. AVARNS000635 sits on 162 catalog
        products and does not exist in RQ at all; every one of those can never show stock. */
  const rqSkus = new Set();
  let rqAsOf = '';
  try {
    const file = JSON.parse(await graph.fileText(STOCK_FILE) || '{}');
    rqAsOf = String(file.asOf || '').slice(0, 10);
    for (const r of file.rows ?? []) if (r.sku) rqSkus.add(norm(r.sku));
  } catch (e) {
    console.log(`  (RepairQ report not read: ${e.message.slice(0, 120)} -- the dead-SKU and inventory checks are skipped)`);
  }

  const products = [...cat.values()];
  console.log(`\ncatalog report: ${products.length} products` +
    (rqSkus.size ? `, ${rqSkus.size} SKUs in the RepairQ report as of ${rqAsOf}` : ', no RepairQ report'));

  // What a catalog product actually carries. Printed because SI's product shape is not
  // documented anywhere we control, and a column we assumed exists would otherwise come
  // out blank for all 5,400 rows without anybody noticing.
  const keys = new Set();
  for (const p of products.slice(0, 200)) for (const k of Object.keys(p.full ?? {})) keys.add(k);
  console.log(`  fields on a catalog product: ${[...keys].sort().join(', ')}`);

  /* The worklist.
     `holds` is deliberately separate from severity. Severity is how much it matters;
     `holds` is whether the product is kept out of the QuickBooks file. They are not the
     same question -- a RepairQ SKU on two products is a real problem that breaks stock
     matching, but the QuickBooks SKU is the Model, so it has no bearing on the import and
     holding the product back for it would be wrong. */
  const issues = [];
  const add = (sev, issue, p, detail, action, holds = false) =>
    issues.push({ sev, issue, model: p?.model ?? '', vendor: p?.vendor ?? '',
                  mfr: mfrOf(p), detail, action, holds });
  const holdModels = new Set();
  const hold = p => { if (p?.model) holdModels.add(p.model.trim().toLowerCase()); };

  /* ---------- the Model, which is the QuickBooks SKU ---------- */

  // Two Models that are the same string once punctuation and spacing are ignored. Usually
  // one product entered twice: ES550TSNDBR50B and ES-550T-SNDBR-50-B. QuickBooks would
  // accept both, which is the problem -- you get two items for one product and the
  // invoicing splits across them.
  //
  // It also catches pairs that only LOOK identical, because stripping punctuation throws
  // away decimal points: "AC-BT-10KUHD .3M" and "AC-BT-10KUHD 3M" are a 0.3m and a 3m
  // cable, two real products. So this says review, not delete. Those pairs are still worth
  // seeing -- a Model where a dropped dot changes which product you get is a Model that
  // will be mis-scanned sooner or later.
  const byNorm = group(products, p => norm(p.model));
  for (const g of byNorm) {
    add('blocker', 'Two products, same Model ignoring punctuation', g[0],
      g.map(p => `"${p.model}"`).join(' vs '),
      'If they are one product, merge and keep one Model. If they are genuinely different, make the Models differ by more than punctuation', true);
    g.forEach(hold);
  }

  // QuickBooks treats a SKU as unique without regard to case, so these collide on import
  // even though D-Tools is happy holding both.
  const byCase = group(products, p => p.model.trim().toLowerCase());
  for (const g of byCase) {
    add('blocker', 'Models differ only by case', g[0], g.map(p => `"${p.model}"`).join(' vs '),
      'QuickBooks cannot hold both -- pick one spelling', true);
    g.forEach(hold);
  }

  for (const p of products) {
    const m = p.model;
    // A colon is how QuickBooks separates a sub-item from its parent, so it cannot appear
    // in a name. Over-length is refused outright. Both stop the import dead.
    if (/:/.test(m) || m.length > QBO_MAX) {
      add('blocker', 'Model QuickBooks will not accept', p,
        `"${m}" -- ${/:/.test(m) ? 'contains a colon, which QuickBooks reserves for sub-items' : `${m.length} characters, over the ${QBO_MAX} allowed`}`,
        'Retype the Model as the manufacturer prints the part number', true);
      hold(p);
    }
    // These do not break the import -- the CSV writer escapes them properly -- but a quote
    // or a double space in a part number is a scan that will not match one day.
    const messy = [];
    if (m !== m.trim()) messy.push('leading or trailing space');
    if (/\s{2,}/.test(m)) messy.push('double space');
    if (/["']/.test(m)) messy.push('quote character');
    if (messy.length) {
      add('warn', 'Model is untidy', p, `"${m}" -- ${messy.join(', ')}`,
        'Retype it cleanly -- this is the string an installer scans');
    }
  }

  /* ---------- the RepairQ SKU ----------
     Nothing here holds the import. These break stock matching in D-Tools and on the Orders
     page, which is a separate job from getting items into QuickBooks. */

  // On more than one product. catalogVendors already blanks these so a wrong stock figure
  // is never shown -- which is why the Orders page shows "?" instead. Correct, but it only
  // clears with a catalog edit. These are those edits.
  const sharedSkus = new Set();
  for (const s of cat.sharedRq ?? []) {
    sharedSkus.add(norm(s.sku));
    add('blocker', 'RepairQ SKU on more than one product', { model: s.models[0] },
      `${s.sku} is on ${s.models.length}: ${s.models.slice(0, 6).join(' / ')}${s.models.length > 6 ? ` / +${s.models.length - 6} more` : ''}`,
      'Leave the SKU on the one product it belongs to, clear CustomField1 on the rest');
  }

  /* A SKU that does not appear in the stock report Brian loaded.
     This is information, NOT a problem, and it was a blocker in the first version of this
     file by mistake. The report is a STOCK report: an item sitting at zero with nothing
     committed and nothing on order has no reason to be in it. When this check first ran it
     produced 694 rows, every one a different SKU on a single product, all of them in
     RepairQ's own format -- which is what a list of real items at zero looks like, not a
     list of typos.
     Telling a genuinely wrong SKU from a correct one at zero stock needs a full RepairQ
     ITEM export to compare against, which is not what gets loaded on the Orders page. Until
     that exists this says only what it can honestly say: no stock figure is available. */
  if (rqSkus.size) {
    const noFigure = [];
    for (const p of products) {
      const raw = rqRaw(p);
      if (!raw || rqSkus.has(norm(raw)) || sharedSkus.has(norm(raw))) continue;
      noFigure.push(p);
    }
    for (const p of noFigure) {
      add('info', 'No stock figure for this SKU', p,
        `${rqRaw(p)} is not in the stock report of ${rqAsOf || 'unknown date'}`,
        'Usually means zero stock with nothing on order, which is fine. Only worth checking if you expected stock on this product');
    }
    if (noFigure.length) {
      console.log(`  (${noFigure.length} products carry a RepairQ SKU that is not in the stock report -- ` +
        'reported as information, not a problem: a stock report has no reason to list an item at zero)');
    }
  }

  // Not a SKU at all -- someone typed a quantity or a dash into the field.
  for (const p of products) {
    const raw = rqRaw(p);
    if (!raw) continue;
    if (/^[\d.,\s]+$/.test(raw) || raw.replace(/[^A-Za-z0-9]/g, '').length < 4) {
      add('blocker', 'RepairQ SKU is not a SKU', p, `CustomField1 reads "${raw}"`,
        'Clear it, or put the real RQ SKU in');
    }
  }

  /* ---------- the fields a new product gets created without ----------
     None of these stop an import. They are what makes ordering, costing and quoting work. */

  for (const p of products) {
    const svc = isService(p);
    if (!svc && !p.vendor && !p.doNotOrder) {
      add('warn', 'No vendor', p, 'Nothing in the Vendor field',
        'Set it in the catalog -- one edit fixes every job the product is on, and it becomes the Preferred Vendor in QuickBooks');
    }
    if (!svc && !(p.cost > 0) && !p.doNotOrder) {
      add('warn', 'No cost', p, 'UnitCost is zero or blank', 'Put the buy price in');
    }
    if (!svc && !mfrOf(p)) {
      add('warn', 'No manufacturer', p, 'Manufacturer is blank', 'Fill it in -- it is how products group when quoting');
    }
    if (!categoryOf(p)) {
      add('warn', 'No category', p, 'Category is blank',
        'Set it -- it carries into QuickBooks as the item category and is how the item list stays navigable');
    }
    const desc = descOf(p);
    if (!desc) {
      add('warn', 'No description', p, 'Description is blank',
        'Write what the thing is -- this text goes on the customer invoice');
    } else if (norm(desc) === norm(p.model)) {
      add('info', 'Description is just the Model', p, `"${desc}"`,
        'Write what it is -- a part number on an invoice means nothing to a customer');
    }
    if (p.doNotOrder) {
      add('info', 'Do Not Order is set', p, 'Excluded from ordering on every job',
        'Correct for labour and allowances -- worth checking it is deliberate on a real product');
    }
    if (isDiscontinued(p)) {
      add('info', 'Discontinued', p, 'Flagged discontinued in SI',
        'Left out of the QuickBooks import -- no action needed unless it is still being bought');
    }
  }

  /* ---------- the files ---------- */

  mkdirSync(OUT, { recursive: true });

  const rank = { blocker: 0, warn: 1, info: 2 };
  issues.sort((a, b) => rank[a.sev] - rank[b.sev] || a.issue.localeCompare(b.issue) || a.model.localeCompare(b.model));
  writeCsv(`${OUT}/catalog-issues.csv`,
    ['Severity', 'Holds import', 'Issue', 'Model', 'Manufacturer', 'Vendor', 'Detail', 'What to do'],
    issues.map(i => [i.sev, i.holds ? 'yes' : '', i.issue, i.model, i.mfr, i.vendor, i.detail, i.action]));

  const issueCount = new Map();
  for (const i of issues) {
    const k = i.model.trim().toLowerCase();
    if (k) issueCount.set(k, (issueCount.get(k) ?? 0) + 1);
  }

  const row = p => ({
    model: p.model, mfr: mfrOf(p), desc: descOf(p), vendor: p.vendor, part: p.part,
    category: categoryOf(p), cost: p.cost || 0, price: priceOf(p) || 0,
    rq: rqRaw(p), usable: p.rq,
    inRq: rqSkus.size ? (rqRaw(p) && rqSkus.has(norm(rqRaw(p))) ? 'yes' : 'no') : '',
    doNotOrder: p.doNotOrder ? 'yes' : '', type: qboType(p, rqSkus), id: p.id,
    discontinued: isDiscontinued(p), updated: updatedOf(p),
    // Either spelling counts -- see addUsed() in sync.mjs for why the stripped form matters.
    used: usedModels.has(p.model.trim().toLowerCase()) || usedModels.has(norm(p.model)),
  });

  const all = products.sort((a, b) => a.model.localeCompare(b.model)).map(row);

  writeCsv(`${OUT}/catalog-master.csv`,
    ['Model', 'Manufacturer', 'Description', 'Category', 'Vendor', 'Part number', 'Unit cost',
     'Unit price', 'RepairQ SKU', 'Usable for stock', 'In RepairQ', 'Do not order',
     'Discontinued', 'Used on a job', 'Last edited in SI', 'QuickBooks item type',
     'Problems', 'Import ready', 'D-Tools Id'],
    all.map(r => [r.model, r.mfr, r.desc, r.category, r.vendor, r.part, r.cost || '', r.price || '',
      r.rq, r.usable ? 'yes' : '', r.inRq, r.doNotOrder,
      r.discontinued ? 'yes' : '', r.used ? 'yes' : '', r.updated, r.type,
      issueCount.get(r.model.trim().toLowerCase()) ?? 0,
      importable(r, holdModels) ? 'yes' : 'no', r.id]));

  /* Suppliers first, items second. Preferred Vendor only resolves if the vendor already
     exists in QuickBooks -- import items first and the column is dropped on every row
     without saying so, and 57 suppliers then have to be set by hand on 5,000 items. */
  const suppliers = [...new Set(all.map(r => r.vendor).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  writeCsv(`${OUT}/qbo-suppliers.csv`, ['Vendor name'], suppliers.map(s => [s]));

  const ready = all.filter(r => importable(r, holdModels));
  const files = [];
  for (let i = 0; i < ready.length; i += QBO_ROWS) {
    const part = ready.slice(i, i + QBO_ROWS);
    const name = `${OUT}/qbo-items-${Math.floor(i / QBO_ROWS) + 1}.csv`;
    writeCsv(name, QBO_COLUMNS.map(([h]) => h), part.map(r => QBO_COLUMNS.map(([, f]) => f(r))));
    files.push(`${name} (${part.length})`);
  }

  /* ---------- what it found ---------- */

  const byIssue = new Map();
  for (const i of issues) {
    const k = `${i.sev}|${i.issue}`;
    byIssue.set(k, (byIssue.get(k) ?? 0) + 1);
  }
  console.log('\n  what needs fixing, worst first:');
  for (const [k, n] of [...byIssue].sort((a, b) => rank[a[0].split('|')[0]] - rank[b[0].split('|')[0]] || b[1] - a[1])) {
    const [sev, issue] = k.split('|');
    console.log(`    ${String(n).padStart(5)}  ${sev.padEnd(7)}  ${issue}`);
  }

  const types = new Map();
  for (const r of ready) types.set(r.type, (types.get(r.type) ?? 0) + 1);
  // Said out loud, because a filter that quietly halves the item count is a filter somebody
  // will later assume was a bug.
  const skipped = all.filter(r => !importable(r, holdModels) &&
    !holdModels.has(r.model.trim().toLowerCase())).length;
  console.log(
    `\n  left out of the QuickBooks file:` +
    `\n    ${holdModels.size} with a problem to fix first` +
    `\n    ${skipped} neither stocked nor ever quoted` +
    `\n  (${all.filter(r => r.discontinued && importable(r, holdModels)).length} of those going in are discontinued` +
    ' -- kept because they are stocked or have been quoted)' +
    `\n  ${ready.length} going in: ` + [...types].sort((a, b) => b[1] - a[1]).map(([t, n]) => `${n} ${t}`).join(', ') +
    `\n  ${suppliers.length} suppliers, to import as vendors BEFORE the items` +
    (QTY_AS_OF ? `\n  Inventory rows carry quantity 0 as of ${QTY_AS_OF}` :
                 '\n  Inventory rows carry no quantity (set CATALOG_QTY_AS_OF=MM/DD/YYYY for the live import)'));
  for (const f of files) console.log(`    ${f}`);
  console.log(`    ${OUT}/qbo-suppliers.csv (${suppliers.length})`);
  console.log(`    ${OUT}/catalog-issues.csv (${issues.length} rows)  <- the worklist`);
  console.log('  Download these from the Artifacts section at the bottom of this run page.');
}

/* ---------- helpers ----------
   SI publishes a product with whatever fields that product has, so a field can be absent
   rather than empty. Each of these tries the names SI is known to use and falls back to
   blank, instead of writing "undefined" into a spreadsheet. The field list printed at the
   top of the run says which names actually exist, so these can be narrowed once we know. */

function mfrOf(p) { const f = p?.full ?? {}; return str(f.Manufacturer || f.Mfr); }
function descOf(p) { const f = p?.full ?? {}; return str(f.Description || f.ShortDescription || f.LongDescription); }
function priceOf(p) { const f = p?.full ?? {}; return Number(f.UnitPrice ?? f.Price) || 0; }
function rqRaw(p) { const v = str(p?.full?.CustomField1); return (v && v !== '-') ? v : ''; }
/* When SI last published a change to this product. Worth having in the export for two
   reasons: it tells you whether a fix you remember making actually reached the feed, and
   it separates products in live use from ones untouched for years. Sliced, not parsed --
   D-Tools sends local wall-clock time with no zone, and constructing a Date here would
   shift the day. */
function updatedOf(p) { return str(p?.full?.UpdatedOn).slice(0, 10); }

/* Category, as "Category:Subcategory" -- QuickBooks reads the colon as one level of
   nesting, per its own sample file. Any colon already inside either part is stripped,
   because two colons would make a sub-sub-category QuickBooks does not support. */
function categoryOf(p) {
  const f = p?.full ?? {};
  const parts = [str(f.Category), str(f.SubCategory || f.Subcategory)]
    .map(s => s.replace(/:/g, ' ').replace(/\s+/g, ' ').trim())
    .filter(Boolean);
  return parts.join(':');
}

/* Labour and allowances, which are services in QuickBooks and have no vendor or cost by
   design.
   A catalog product carries no TypeId -- the field simply is not in SI's product payload,
   as the field list printed at the top of the run shows -- so labour is identified from
   what is there: labour hours or labour types on the product, or a category that says so.
   Do Not Order is the last fallback, being a flag somebody set deliberately.

   A blank vendor and no cost is NOT used as a signal here, however much it looks like one.
   That is also exactly what a carelessly created real product looks like, and treating
   those as services would quietly excuse them from the checks meant to catch them. */
const SERVICE_CATEGORY = /^(labou?r|services?|per ?diem|levy|freight|shipping|subcontract|allowance)/i;
function isService(p) {
  const f = p?.full ?? {};
  return SERVICE_CATEGORY.test(str(f.Category)) || SERVICE_CATEGORY.test(str(f.Subcategory));
}

// SI has a Discontinued flag and the catalog carries plenty of them. They are history --
// worth keeping in D-Tools so old quotes still read, worth keeping out of QuickBooks,
// where every one is a row in the item list somebody has to scroll past for ever.
function isDiscontinued(p) { return !!(p?.full ?? {}).Discontinued; }

/* Inventory means QuickBooks keeps a quantity and FIFO cost layers for it, and somebody
   counts it at year end. Worth doing for what sits on the shelf, not for gear bought per
   job -- and RepairQ already knows which is which, so it decides rather than us.
   The raw SKU is used, not the blanked one: a SKU shared between two products still tells
   us the thing is stocked, even though it cannot tell us how many. */
function qboType(p, rqSkus) {
  if (isService(p)) return 'Service';
  const raw = rqRaw(p);
  if (raw && rqSkus.has(norm(raw))) return 'Inventory';
  return 'Non-Inventory';
}

/* What goes into the QuickBooks file: stocked, quoted, or labour. Everything else stays in
   D-Tools, where it costs nothing. If a never-quoted product gets specced in 2027, the
   estimate push creates the item then -- nothing is lost by leaving it out now, and every
   row left out is one nobody scrolls past and nobody has to clean.

   Labour is in regardless of whether it looks quoted, and that is deliberate. orderLines()
   filters to buyable lines -- it drops TypeId 2, Do Not Order and anything with no cost --
   so labour can never appear in usedModels however many jobs carry it. Labour and
   allowances are exactly what an invoice needs, so excluding them would break the estimate
   push the moment a job had a labour line. It is a small set either way. */
function importable(r, holdModels) {
  if (holdModels.has(r.model.trim().toLowerCase())) return false;
  /* Discontinued is deliberately NOT a reason on its own.
     A discontinued product that is still stocked needs an item for the opening inventory
     count, and one that has been quoted needs an item to invoice against -- a job
     straddling 31 December with a discontinued part on it would otherwise have nothing to
     bill. Discontinued products nobody stocks or has quoted are already left out by the
     rule below, so checking the flag separately only ever excluded the wrong ones. */
  return r.type === 'Inventory' || r.type === 'Service' || r.used;
}

// Groups sharing a key, groups of one dropped -- the shape every duplicate check wants.
function group(items, keyOf) {
  const m = new Map();
  for (const it of items) {
    const k = keyOf(it);
    if (!k) continue;
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(it);
  }
  return [...m.values()].filter(g => g.length > 1);
}

function str(v) { return String(v ?? '').trim(); }
function norm(s) { return String(s ?? '').replace(/[^A-Za-z0-9]/g, '').toUpperCase(); }

// Excel opens a CSV in the system encoding unless there is a byte order mark, which turns
// every accented supplier name into mojibake. The BOM is three bytes and prevents that.
function writeCsv(path, head, rows) {
  const cell = v => {
    const s = String(v ?? '');
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const body = [head, ...rows].map(r => r.map(cell).join(',')).join('\r\n');
  writeFileSync(path, '﻿' + body + '\r\n', 'utf8');
}
