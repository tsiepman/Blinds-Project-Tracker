// D-Tools → SharePoint sync for the project board.
//
//   node sync.mjs            normal run
//   node sync.mjs --dry-run  report what would change, write nothing
//
// Runs on a schedule in GitHub Actions. The wall board reads the SharePoint list it
// writes; nothing in the browser ever sees the D-Tools key.

import { DTools, phaseEstimate, taskSpans, orderLines, catalogVendors,
         visitFromTask, visitFromServiceOrder, visitInWindow } from './dtools.mjs';
import { Graph } from './graph.mjs';
import { processWriteups } from './writeups.mjs';
import { pushStock } from './stockpush.mjs';
import { createHash } from 'node:crypto';

const SITE_HOSTNAME = 'advanceelectronics.sharepoint.com';
const SITE_PATH = '/sites/ART43';
const LIST = 'SchedProjects';
// The Orders page's small shared list. Its "vendors" row holds every supplier name in the
// catalog, for the lead-time picker; the page itself owns the "stock" row.
const STOCK_LIST = 'StockSnapshot';
// Who was scheduled where, written down as it happens -- see "field visits" in dtools.mjs.
// Rows are never deleted here: a visit that disappears from D-Tools is exactly the thing
// worth still having a record of.
const VISIT_LIST = 'FieldVisits';

// Every approved project is synced, whatever its price. The Orders page reads this list and
// needs every job that has product to buy; the wall board applies its own $10k cut when it
// reads (MIN_PRICE in ART_Project_Board.html). Filtering here once hid 46 jobs from ordering.

const dryRun = process.argv.includes('--dry-run');

async function main() {
  const si = new DTools(process.env.DTOOLS_API_KEY);
  const graph = new Graph({
    tenantId: process.env.AZURE_TENANT_ID,
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
  });

  // D-Tools first, SharePoint second. The D-Tools half needs nothing but the API key, so
  // a dry run can prove it works -- endpoints, catalog map, parsing -- before the Azure
  // admin consent exists. SharePoint is only touched once there is something to write.
  const projects = await si.projects();
  console.log(`Projects published to the integration: ${projects.length}`);

  // Vendor is resolved against the catalog rather than the project item, so one catalog
  // fix reaches every job instead of needing an item-by-item edit in each one.
  const catalogs = (await si.listCatalogs()).Catalogs;
  console.log(`Catalog messages in the feed: ${catalogs.length}`);

  // Fingerprint of the catalog feed, stored in CatalogSeen. Every catalog edit arrives as
  // a new message, so the list of message ids moves whenever a vendor fix does -- and
  // reading that list is one small request, where reading the products means fetching
  // every message (the base catalog alone is ~9,000 product rows).
  const catalogPrint = createHash('sha1')
    .update(catalogs.map(c => c.Id).join('\n'))
    .digest('hex').slice(0, 16);

  // So products are only read when some project actually needs rebuilding, and at most
  // once per run. On a quiet twenty-minute tick nothing needs it.
  let vendorMap = null;
  const getVendorMap = async () => {
    if (!vendorMap) {
      vendorMap = await catalogVendors(si, catalogs);
      console.log(`  (catalog read: ${vendorMap.size} products)`);
      // A RepairQ SKU on two products can't identify either, so both lose their stock
      // figure on the Orders page. Each line here is one catalog edit that fixes it.
      const shared = vendorMap.sharedRq ?? [];
      if (shared.length) {
        console.log(`  ${shared.length} RepairQ SKU${shared.length === 1 ? '' : 's'} on more than one product (stock can't be matched for these):`);
        for (const s of shared.slice(0, 15)) console.log(`    ${s.sku}  ${s.models.slice(0, 4).join(' · ')}${s.models.length > 4 ? ` · +${s.models.length - 4} more` : ''}`);
      }
    }
    return vendorMap;
  };

  const existing = new Map();
  let vendorsRow = null, dupesRow = null;
  try {
    await graph.siteId(SITE_HOSTNAME, SITE_PATH);
    for (const it of await graph.items(LIST)) {
      const f = it.fields ?? {};
      if (f.ProjectId) existing.set(f.ProjectId, { itemId: it.id, ...f });
    }
    console.log(`Already in SchedProjects: ${existing.size}`);
    for (const it of await graph.items(STOCK_LIST)) {
      const t = (it.fields ?? {}).Title;
      if (t === 'vendors') vendorsRow = it;
      else if (t === 'dupes') dupesRow = it;
    }
  } catch (e) {
    // A real run must not continue blind -- it would create duplicates of everything.
    if (!dryRun) throw e;
    console.log(`\nSharePoint not reachable (${e.message.slice(0, 120)})`);
    console.log('Dry run continues as if the list were empty. Expected until Azure admin consent is granted.\n');
  }

  let created = 0, updated = 0, unchanged = 0;
  const failed = [];
  const visits = [];
  const stats = { withStart: 0, withTasks: 0, board: 0, active: 0, lines: 0, catalog: 0, project: 0, noVendor: 0 };
  const noVendorModels = new Map();                 // model -> number of jobs it is on

  for (const p of projects) {
    const prior = existing.get(p.Id);

    // The full project record runs to well over a megabyte because it carries every
    // line item. Only refetch when SI says the project actually changed -- otherwise a
    // twenty-minute cron would pull tens of megabytes an hour to learn nothing.
    const changed = !prior || String(prior.SourceUpdatedOn || '') !== String(p.UpdatedOn || '');
    // A catalog vendor fix changes no project, so order lines would go stale if we only
    // watched the project. Rebuild them when the catalog feed has moved as well.
    const catalogMoved = !prior || String(prior.CatalogSeen || '') !== catalogPrint;
    const rebuild = changed || catalogMoved;

    let estimate = prior
      ? { hours: { roughIn: num(prior.HoursRoughIn), finish: num(prior.HoursFinish) },
          days:  { roughIn: num(prior.DaysRoughIn),  finish: num(prior.DaysFinish) } }
      : null;
    let builder = prior?.Builder ?? '';
    let startDate = prior?.StartDate ?? '';
    let orders = prior?.OrderLinesJson ?? '[]';

    if (rebuild) {
      const detail = await si.projectDetail(p.Id);
      estimate = phaseEstimate(detail);
      // CustomField1 on the project holds the builder -- "Maric Homes" on the job this
      // was built against. It is what the contact tracker keys on.
      builder = detail.CustomField1 || '';
      startDate = dayOrEmpty(detail.StartDate);
      // Filtered to buyable lines before storing. A job carries ~350 items; forty jobs
      // would be 14,000 rows, past SharePoint's 5,000-item query threshold. Filtering
      // first cuts a job to roughly 70, so this stays one row per project and the
      // Orders page aggregates across them in a single request.
      orders = JSON.stringify(orderLines(detail, await getVendorMap()));
    }

    // Tasks change independently of the project, so refresh them every run -- but only for
    // jobs still open. A Completed job's schedule is history; it is read once, the first
    // time the job is seen, and kept. Otherwise every finished job costs a request every
    // twenty minutes forever. (Reopening a job in D-Tools changes its Progress, so it
    // starts refreshing again.)
    const active = !String(p.Progress || '').trim().toLowerCase().startsWith('complete');
    let spans;
    if (active || !prior) {
      const tasks = await si.tasks(p.Id);
      spans = taskSpans(tasks);
      // The same fetch feeds the field-visit log: who was scheduled where, recorded while
      // the task still exists.
      for (const t of tasks ?? []) {
        const v = visitFromTask(t, p);
        if (v && visitInWindow(v)) visits.push(v);
      }
    }
    else { try { spans = JSON.parse(prior.TasksJson || '[]'); } catch { spans = []; } }

    const fields = {
      Title: p.Number || p.Name || p.Id,
      ProjectId: p.Id,
      ProjectName: p.Name || '',
      ProjectNumber: p.Number || '',
      Client: p.Client || '',
      Builder: builder,
      Progress: p.Progress || '',
      StartDate: startDate,
      Price: String(p.Price ?? ''),
      HoursRoughIn: String(estimate.hours.roughIn),
      HoursFinish: String(estimate.hours.finish),
      DaysRoughIn: String(estimate.days.roughIn),
      DaysFinish: String(estimate.days.finish),
      // Denormalised on purpose. The board wants one request, not a join across two
      // lists, and this is a cache rather than a database.
      TasksJson: JSON.stringify(spans),
      TaskCount: String(spans.length),
      OrderLinesJson: orders,
      SourceUpdatedOn: p.UpdatedOn || '',
      CatalogSeen: catalogPrint,
      SyncedAt: new Date().toISOString(),
    };

    // Tallies for the end-of-run report.
    const lines = JSON.parse(orders);
    stats.withStart += startDate ? 1 : 0;
    stats.withTasks += spans.length ? 1 : 0;
    stats.board += Number(p.Price || 0) >= 10000 ? 1 : 0;   // report only -- matches the board's MIN_PRICE
    // Order-line figures count only jobs the Orders page shows. It hides Completed jobs,
    // and a missing vendor on a finished job is not worth anyone's time to fix.
    stats.active += active ? 1 : 0;
    if (active) stats.lines += lines.length;
    for (const l of active ? lines : []) {
      stats[l.vendorSource || 'noVendor']++;
      if (!l.vendorSource) noVendorModels.set(l.model, (noVendorModels.get(l.model) ?? 0) + 1);
    }

    // One project SharePoint rejects must not stop the rest being written. Log it, carry
    // on, and fail the run at the end so GitHub still flags it. A rerun is safe: rows are
    // matched on ProjectId, so anything already written is updated, not duplicated.
    try {
      if (!prior) {
        if (!dryRun) await graph.create(LIST, fields);
        created++;
        console.log(`  + ${String(fields.Title).padEnd(11)} start ${(startDate || '—').padEnd(10)}  tasks ${String(spans.length).padStart(3)}  lines ${String(lines.length).padStart(3)}  ${fields.ProjectName}`);
      } else if (rebuild || String(prior.TasksJson || '') !== fields.TasksJson) {
        if (!dryRun) await graph.update(LIST, prior.itemId, fields);
        updated++;
        console.log(`  ~ ${fields.Title} — ${spans.length} task(s)`);
      } else {
        unchanged++;
      }
    } catch (e) {
      failed.push(fields.Title);
      console.log(`  ! ${fields.Title} not saved: ${e.message.slice(0, 200)}`);
    }
  }

  /* ---------- field visits ----------
     Service calls come from their own feed; project work came out of the task fetch above.
     Rows are matched on the D-Tools id, updated while the visit still exists, and left
     alone once it stops appearing -- a deleted or rescheduled visit keeps its record. */
  try {
    for (const s of await si.serviceOrders()) {
      const v = visitFromServiceOrder(await si.getServiceOrder(s.Id));
      if (v && visitInWindow(v)) visits.push(v);
    }
  } catch (e) {
    console.log(`  ! service orders not read: ${e.message.slice(0, 160)}`);
  }

  let visitsNew = 0, visitsChanged = 0;
  if (visits.length) {
    try {
      const have = new Map();
      for (const it of await graph.items(VISIT_LIST)) {
        const f = it.fields ?? {};
        if (f.VisitId) have.set(f.VisitId, { itemId: it.id, ...f });
      }
      for (const v of visits) {
        const prior = have.get(v.visitId);
        const fields = {
          Title: `${v.date} ${v.jobNumber || v.client || v.name}`.slice(0, 250),
          VisitId: v.visitId,
          Kind: v.kind,
          Number: v.number,
          JobNumber: v.jobNumber,
          JobName: String(v.jobName).slice(0, 250),
          Client: String(v.client).slice(0, 250),
          Site: String(v.site).slice(0, 250),
          VisitName: String(v.name).slice(0, 250),
          SchedDate: v.date,
          SchedFrom: v.from,
          SchedTo: v.to,
          Crew: v.crew.join(', ').slice(0, 250),
          Progress: v.progress,
          PercentComplete: String(v.pct),
          Instructions: v.instructions,
          SourceUpdatedOn: v.updatedOn,
          SourceUpdatedBy: v.updatedBy,
          SyncedAt: new Date().toISOString(),
        };
        try {
          if (!prior) { if (!dryRun) await graph.create(VISIT_LIST, { ...fields, FirstSeen: new Date().toISOString() }); visitsNew++; }
          else {
            // Only the fields that can move -- so a re-run doesn't rewrite every row, and
            // the write-up columns the field log fills in later are never touched.
            const moved = ['SchedDate', 'SchedFrom', 'SchedTo', 'Crew', 'Progress', 'PercentComplete', 'Instructions', 'SourceUpdatedOn']
              .some(k => String(prior[k] ?? '') !== String(fields[k] ?? ''));
            if (moved) { if (!dryRun) await graph.update(VISIT_LIST, prior.itemId, fields); visitsChanged++; }
          }
        } catch (e) {
          failed.push('visit ' + v.number);
          console.log(`  ! visit ${v.jobNumber} ${v.date} not saved: ${e.message.slice(0, 160)}`);
        }
      }
      console.log(`\nfield visits: ${visits.length} in the window  ${visitsNew} new  ${visitsChanged} changed`);
    } catch (e) {
      // The list may not exist yet. Say so once; everything else still syncs.
      console.log(`\nField visits not saved -- is the ${VISIT_LIST} list created? (${e.message.slice(0, 120)})`);
    }
  }

  // RepairQ stock back into the catalog's custom fields, so a salesman speccing a job in
  // SI sees what is on the shelf. Only runs when Brian has loaded a newer report.
  try { await pushStock(si, graph, getVendorMap, { dryRun }); }
  catch (e) { failed.push('stock push'); console.log(`  ! stock push: ${e.message.slice(0, 200)}`); }

  // Installer write-ups: notes the Power Automate flow copied into WorkOrderNotes are
  // summarised once and attached to their field visit. See writeups.mjs.
  try { await processWriteups(graph, { dryRun }); }
  catch (e) { failed.push('write-ups'); console.log(`  ! write-ups: ${e.message.slice(0, 200)}`); }

  // Supplier names for the lead-time picker, refreshed only when the catalog feed has moved.
  // AsOf carries the catalog fingerprint on this row, not a date, so the check is one compare.
  if (!vendorsRow || String(vendorsRow.fields?.AsOf || '') !== catalogPrint) {
    const names = [...new Set([...(await getVendorMap()).values()].map(v => v.vendor).filter(Boolean))]
      .sort((a, b) => a.localeCompare(b));
    const fields = { Title: 'vendors', DataJson: JSON.stringify(names), AsOf: catalogPrint };
    try {
      if (!dryRun) {
        if (vendorsRow) await graph.update(STOCK_LIST, vendorsRow.id, fields);
        else await graph.create(STOCK_LIST, fields);
      }
      console.log(`\nsuppliers in the catalog: ${names.length}${dryRun ? '' : ' (saved for the lead-time picker)'}`);
    } catch (e) {
      failed.push('supplier list');
      console.log(`  ! supplier list not saved: ${e.message.slice(0, 200)}`);
    }

    // The RepairQ SKUs sitting on more than one product, so the Orders page can name the
    // clash on the product it affects instead of leaving it as an unexplained "?".
    const shared = (await getVendorMap()).sharedRq ?? [];
    const dupeFields = { Title: 'dupes', DataJson: JSON.stringify(shared).slice(0, 60000), AsOf: catalogPrint };
    try {
      if (!dryRun) {
        if (dupesRow) await graph.update(STOCK_LIST, dupesRow.id, dupeFields);
        else await graph.create(STOCK_LIST, dupeFields);
      }
    } catch (e) {
      failed.push('duplicate SKU list');
      console.log(`  ! duplicate SKU list not saved: ${e.message.slice(0, 200)}`);
    }
  }

  console.log(`\ncreated ${created}  updated ${updated}  unchanged ${unchanged}`);
  const shown = created + updated + unchanged;
  console.log(
    `\nprojects synced           : ${shown}  (all go to the Orders page)` +
    `\n  $10k and over           : ${stats.board}  (the wall board shows these)` +
    `\nprojects with a StartDate : ${stats.withStart} of ${shown}` +
    `\nprojects with tasks       : ${stats.withTasks} of ${shown}` +
    `\nnot Completed             : ${stats.active}  (what the Orders page shows; figures below are these jobs only)` +
    `\norder lines               : ${stats.lines}` +
    `\n  vendor from catalog     : ${stats.catalog}` +
    `\n  vendor from project copy: ${stats.project}` +
    `\n  no vendor at all        : ${stats.noVendor}`
  );
  // The catalog fixes worth doing first: products with no vendor, by how many jobs they
  // are on. A fix in the catalog reaches every one of those jobs on the next run.
  if (noVendorModels.size) {
    console.log(`\nno vendor on jobs not yet Completed, most jobs first (${noVendorModels.size} products):`);
    [...noVendorModels].sort((a, b) => b[1] - a[1]).slice(0, 25)
      .forEach(([m, n]) => console.log(`  ${String(n).padStart(3)}  ${m}`));
  }
  if (dryRun) console.log('--dry-run: nothing was written.');
  if (failed.length) {
    console.error(`\n${failed.length} project(s) not saved: ${failed.join(', ')}`);
    process.exitCode = 1;
  }
  if (!projects.length) {
    console.log(
      '\nNo projects in the queue. The Subscribe feed is delta-based, so a project\n' +
      'appears only once it has been edited in SI since publishing was enabled.'
    );
  }
}

function num(v) { return Number(v) || 0; }
// D-Tools returns local wall-clock timestamps with no zone. Slice rather than parse --
// constructing a Date here would shift dates by one across timezones.
// D-Tools also uses 0001-01-01 as an "empty" placeholder (tasks arrive with it), which
// would otherwise be written as a real date two thousand years ago.
function dayOrEmpty(v) {
  const s = v ? String(v).slice(0, 10) : '';
  return s.startsWith('0001-') ? '' : s;
}

main().catch(e => { console.error('\n' + e.message); process.exit(1); });
