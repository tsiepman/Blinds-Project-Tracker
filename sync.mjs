// D-Tools → SharePoint sync for the project board.
//
//   node sync.mjs            normal run
//   node sync.mjs --dry-run  report what would change, write nothing
//
// Runs on a schedule in GitHub Actions. The wall board reads the SharePoint list it
// writes; nothing in the browser ever sees the D-Tools key.

import { DTools, phaseEstimate, taskSpans, orderLines, catalogVendors } from './dtools.mjs';
import { Graph } from './graph.mjs';
import { createHash } from 'node:crypto';

const SITE_HOSTNAME = 'advanceelectronics.sharepoint.com';
const SITE_PATH = '/sites/ART43';
const LIST = 'SchedProjects';

// Projects below this price are left off the board. A wall display is readable at about
// thirty bars and useless at three hundred, so this is the volume control. Set as a
// repository variable rather than edited here.
const MIN_PRICE = Number(process.env.MIN_PRICE || 0);

const dryRun = process.argv.includes('--dry-run');

async function main() {
  const si = new DTools(process.env.DTOOLS_API_KEY);
  const graph = new Graph({
    tenantId: process.env.AZURE_TENANT_ID,
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
  });

  await graph.siteId(SITE_HOSTNAME, SITE_PATH);

  const projects = await si.projects();
  console.log(`Projects published to the integration: ${projects.length}`);

  // Vendor is resolved against the catalog rather than the project item, so one catalog
  // fix reaches every job instead of needing an item-by-item edit in each one.
  const vendorMap = await catalogVendors(si);
  console.log(`Catalog products seen in the feed: ${vendorMap.size}`);

  // A fingerprint of everything order lines take from the catalog. A vendor fix on an
  // existing product leaves the product count alone, so counting would miss it; this
  // changes on any edit that matters. Stored in CatalogSeen.
  const catalogPrint = createHash('sha1')
    .update([...vendorMap.entries()].sort(([a], [b]) => a.localeCompare(b))
      .map(([k, v]) => [k, v.vendor, v.cost, v.rq, v.part].join('|')).join('\n'))
    .digest('hex').slice(0, 16);

  const existing = new Map();
  for (const it of await graph.items(LIST)) {
    const f = it.fields ?? {};
    if (f.ProjectId) existing.set(f.ProjectId, { itemId: it.id, ...f });
  }
  console.log(`Already on the board: ${existing.size}`);

  let created = 0, updated = 0, skipped = 0, unchanged = 0;

  for (const p of projects) {
    if (MIN_PRICE && Number(p.Price || 0) < MIN_PRICE) { skipped++; continue; }

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
      orders = JSON.stringify(orderLines(detail, vendorMap));
    }

    // Tasks are cheap and change independently of the project, so always refresh them.
    const spans = taskSpans(await si.tasks(p.Id));

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

    if (!prior) {
      if (!dryRun) await graph.create(LIST, fields);
      created++;
      console.log(`  + ${fields.Title} — ${fields.ProjectName}`);
    } else if (rebuild || String(prior.TasksJson || '') !== fields.TasksJson) {
      if (!dryRun) await graph.update(LIST, prior.itemId, fields);
      updated++;
      console.log(`  ~ ${fields.Title} — ${spans.length} task(s)`);
    } else {
      unchanged++;
    }
  }

  console.log(`\ncreated ${created}  updated ${updated}  unchanged ${unchanged}  below MIN_PRICE ${skipped}`);
  if (dryRun) console.log('--dry-run: nothing was written.');
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
