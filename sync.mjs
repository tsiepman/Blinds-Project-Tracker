// QuickBooks Online -> D-Tools SI inventory sync.
//
//   node sync.mjs --harvest      Build product-ids.json (model -> SI product Id) from
//                                the daily Subscribe deltas. Run this first, and
//                                periodically, since it is the only way to learn Ids.
//   node sync.mjs --dry-run      Report what would change. Touches nothing.
//   node sync.mjs                Publish changed quantities to SI.
//   node sync.mjs --status <id>  Check whether a published message was imported.
//
// Set USE_STUB=1 to run against fake quantities instead of QuickBooks. That is the
// intended mode until QBO Plus is purchased (~October 2026).

import { readFile, writeFile } from 'node:fs/promises';
import { DTools, normaliseModel } from './dtools.mjs';
import { QuickBooks, stubInventory } from './qbo.mjs';

const ID_MAP_FILE = './product-ids.json';
const STATE_FILE = './state.json';

// SI custom field slots, as configured in Control Panel -> Product/Labor Custom Fields.
// All three are Number-typed slots. These indices are NOT arbitrary: the API fixes a
// data type per slot, and 6/7/8 are the numeric ones Trevor assigned.
const FIELD = { stock: 'CustomField6', committed: 'CustomField7', ordered: 'CustomField8' };

// SI accepts at most 5000 products per publish.
const MAX_BATCH = 5000;

async function main() {
  const args = process.argv.slice(2);
  const dtools = new DTools(process.env.DTOOLS_API_KEY);

  if (args.includes('--harvest')) return harvest(dtools);

  const statusIdx = args.indexOf('--status');
  if (statusIdx !== -1) {
    const id = args[statusIdx + 1];
    if (!id) throw new Error('--status needs a message id');
    console.log(JSON.stringify(await dtools.messageStatus(id), null, 2));
    return;
  }

  // Safety interlock: stub quantities are fabricated. Publishing them would write
  // nonsense stock numbers onto real products that salespeople quote from, and SI
  // gives no undo. Refuse to publish stub data unless it is asked for explicitly.
  const usingStub = process.env.USE_STUB === '1';
  const forced = args.includes('--allow-stub-publish');
  if (usingStub && !forced && !args.includes('--dry-run')) {
    console.warn(
      'USE_STUB=1 — forcing --dry-run so fake quantities cannot reach the live catalog.\n' +
      'Override with --allow-stub-publish only against a catalog you can afford to dirty.\n',
    );
  }

  await sync(dtools, { dryRun: args.includes('--dry-run') || (usingStub && !forced) });
}

async function harvest(dtools) {
  console.log('Harvesting product Ids from SI daily deltas...');
  const { map, collisions } = await dtools.harvestProductIds({
    onProgress: (done, total, size) =>
      process.stdout.write(`\r  catalog ${done}/${total} — ${size} products`),
  });
  console.log();

  // A duplicated model breaks the match key: only one of the products would ever
  // receive quantities, and which one is arbitrary. Surface it loudly — it is a
  // catalog data problem to fix in SI, not something the sync can resolve.
  if (collisions.length) {
    console.warn(`\nWARNING: ${collisions.length} model(s) map to more than one SI product.`);
    console.warn('Only one of each will receive quantities. Fix these in the SI catalog:');
    for (const c of collisions.slice(0, 20)) {
      console.warn(`  ${c.model} — ${c.ids.length} products (${c.ids.join(', ')})`);
    }
    if (collisions.length > 20) console.warn(`  ...and ${collisions.length - 20} more`);
    console.warn('');
  }

  await writeFile(ID_MAP_FILE, JSON.stringify(Object.fromEntries(map), null, 2));
  console.log(`Wrote ${map.size} products to ${ID_MAP_FILE}`);
  console.log(
    '\nNote: deltas only cover products changed in the retained window, so this is a\n' +
    'subset of the full catalog. It grows as more days are observed.',
  );
}

async function sync(dtools, { dryRun }) {
  const idMap = await readJson(ID_MAP_FILE);
  if (!idMap) throw new Error(`${ID_MAP_FILE} not found — run: node sync.mjs --harvest`);

  const known = Object.entries(idMap);
  console.log(`Known SI products: ${known.length}`);

  const inventory = await loadInventory(known.map(([, p]) => p.Model));
  console.log(`QuickBooks inventory items: ${inventory.size}`);

  const previous = (await readJson(STATE_FILE)) ?? {};
  const changed = [];
  const nextState = {};
  let matched = 0;

  for (const [key, product] of known) {
    const qty = inventory.get(key);

    // No QuickBooks item for this model. This is how labour lines and placeholder
    // products ("Shipping Charges", "Existing Flat Panel TV") get excluded — they are
    // not inventory in QBO, so they never receive quantities. No pattern list needed.
    if (!qty) continue;
    matched++;

    const stamp = `${qty.stock}|${qty.committed}|${qty.ordered}`;
    nextState[key] = stamp;
    if (previous[key] === stamp) continue; // unchanged since last run

    changed.push({
      Id: product.Id,
      IntegrationProductId: product.Model,
      Manufacturer: product.Manufacturer,
      Model: product.Model,
      PartNumber: product.PartNumber,
      [FIELD.stock]: qty.stock,
      [FIELD.committed]: qty.committed,
      [FIELD.ordered]: qty.ordered,
    });
  }

  console.log(`Matched to QuickBooks: ${matched}`);
  console.log(`Quantities changed:    ${changed.length}`);

  if (!changed.length) {
    console.log('Nothing to publish.');
    return;
  }

  if (dryRun) {
    console.log('\n--dry-run, publishing nothing. First 10 changes:');
    for (const p of changed.slice(0, 10)) {
      console.log(`  ${p.Model.padEnd(34)} stock=${p[FIELD.stock]} committed=${p[FIELD.committed]} ordered=${p[FIELD.ordered]}`);
    }
    return;
  }

  const stampUtc = new Date().toISOString().slice(0, 16).replace('T', ' ');
  for (let i = 0; i < changed.length; i += MAX_BATCH) {
    const batch = changed.slice(i, i + MAX_BATCH);
    const res = await dtools.publishCatalog({
      Name: `QBO Inventory Sync ${stampUtc} UTC`,
      Description: 'Automated inventory quantities from QuickBooks Online',
      PriceType: 'A',
      IsMetric: false,
      Products: batch,
    });
    console.log(`Published ${batch.length} products — MessageId ${res.MessageId}`);
    console.log(`  confirm later with: node sync.mjs --status ${res.MessageId}`);
  }

  await writeFile(STATE_FILE, JSON.stringify(nextState, null, 2));
  console.log(
    '\nPublished to the queue. SI applies it on its own schedule (Control Panel ->\n' +
    'Manage Integrations, currently 2 hrs), so expect a delay before it is visible.',
  );
}

async function loadInventory(models) {
  if (process.env.USE_STUB === '1') {
    console.log('USE_STUB=1 — using fake quantities, not QuickBooks.');
    return stubInventory(models);
  }

  const required = ['QBO_CLIENT_ID', 'QBO_CLIENT_SECRET', 'QBO_REFRESH_TOKEN', 'QBO_REALM_ID'];
  const missing = required.filter((k) => !process.env[k]);
  if (missing.length) {
    throw new Error(
      `Missing ${missing.join(', ')}.\n` +
      'QuickBooks is not set up yet — run with USE_STUB=1 to exercise the D-Tools half.',
    );
  }

  const qbo = new QuickBooks({
    clientId: process.env.QBO_CLIENT_ID,
    clientSecret: process.env.QBO_CLIENT_SECRET,
    refreshToken: process.env.QBO_REFRESH_TOKEN,
    realmId: process.env.QBO_REALM_ID,
    environment: process.env.QBO_ENV ?? 'production',
    onTokenRefresh: async (token) => {
      // Rotating token. In GitHub Actions the process is thrown away, so this MUST go
      // somewhere durable or the sync dies once the old token ages out. See README.
      console.warn('QBO refresh token rotated — persist it now.');
      await writeFile('./qbo-refresh-token.txt', token);
    },
  });

  const inv = await qbo.getInventory();
  return new Map([...inv].map(([sku, v]) => [normaliseModel(sku), v]));
}

async function readJson(path) {
  try {
    return JSON.parse(await readFile(path, 'utf8'));
  } catch (err) {
    if (err.code === 'ENOENT') return null;
    throw err;
  }
}

main().catch((err) => {
  console.error(`\n${err.message}`);
  if (err.body) console.error(err.body);
  process.exit(1);
});
