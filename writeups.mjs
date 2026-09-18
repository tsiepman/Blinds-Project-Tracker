// Installer write-ups: from the WorkOrderNotes list to the Field Log.
//
//   D-Tools Mobile Install  ->  "Site Notes added" email  ->  workorders@ mailbox
//   ->  Power Automate flow  ->  WorkOrderNotes list  ->  this  ->  FieldVisits + summary
//
// The sync never touches a mailbox. The flow copies each notification into SharePoint,
// and this reads it from there with the permission the sync already has.
//
// Each note is handled once:
//   1. parse it (notes.mjs) -- task/service-order number, installer, the note, its date
//   2. summarise it with Claude -- done, outstanding and why, office actions, what the
//      client or builder said. Stored on the note row (Result), so no note is paid for twice.
//   3. copy it onto the matching FieldVisits row, which is what the Field Log shows.
//
// WorkOrderNotes needs two columns beyond what the flow fills: ProcessedAt (single line)
// and Result (multiple lines, plain text).

import { parseNotice, daysLate } from './notes.mjs';

const NOTES_LIST = 'WorkOrderNotes';
const VISIT_LIST = 'FieldVisits';

// Opus 5 unless the repository variable CLAUDE_MODEL says otherwise -- claude-sonnet-5
// or claude-haiku-4-5 are cheaper for notes this short. Changing it needs no code change.
const MODEL = process.env.CLAUDE_MODEL || 'claude-opus-5';

/* What we ask for, as a strict schema so the answer is always machine-readable. */
const SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['summary', 'completed', 'outstanding', 'returnVisit', 'officeActions', 'clientBuilderNotes'],
  properties: {
    summary: { type: 'string', description: 'One plain sentence: what happened on this visit.' },
    completed: { type: 'array', items: { type: 'string' }, description: 'Work finished on the visit, as short phrases.' },
    outstanding: {
      type: 'array',
      description: 'Work not finished that someone has to come back for.',
      items: {
        type: 'object', additionalProperties: false, required: ['item', 'reason'],
        properties: {
          item: { type: 'string' },
          reason: { type: 'string', description: 'Why it was not finished, if the note says. Empty string if it does not.' },
        },
      },
    },
    returnVisit: { type: 'boolean', description: 'True if the note says or clearly implies another visit is needed.' },
    officeActions: {
      type: 'array', items: { type: 'string' },
      description: 'Things the office or sales must do: order or bring a part, send a quote or details, call the client, keep a promise made to the client.',
    },
    clientBuilderNotes: {
      type: 'array', items: { type: 'string' },
      description: 'Anything the client or builder said, asked for, complained about, or was promised.',
    },
  },
};

const SYSTEM =
  'You read installer write-ups for ART, a residential audio-video and home automation ' +
  'installer in Winnipeg (Control4, Lutron lighting and shades, networking, cameras, ' +
  'security). Each write-up is a technician\'s short field note from one site visit, ' +
  'usually headed with their initials, the date and hours worked.\n\n' +
  'Pull out only what the note says. Do not guess, infer causes, or add advice. Keep every ' +
  'phrase short, as a scheduler would jot it. Leave a list empty rather than padding it. ' +
  'Travel time and hours are not work items. The readers are the scheduler and the ' +
  'salespeople: what matters most is anything unfinished, anything the office has to act ' +
  'on, and anything the client or builder said.';

/* ---------- Claude ---------- */

let client = null;
async function claude() {
  if (client !== null) return client;
  if (!process.env.ANTHROPIC_API_KEY) { client = false; return client; }
  try {
    const { default: Anthropic } = await import('@anthropic-ai/sdk');
    client = new Anthropic();
  } catch (e) {
    console.log(`  ! Claude SDK not available (${e.message.slice(0, 100)}) -- notes saved without summaries`);
    client = false;
  }
  return client;
}

async function summarise(n) {
  const c = await claude();
  if (!c) return { error: 'no API key' };
  const prompt =
    `Installer: ${n.by || 'unknown'}\nClient: ${n.client || '-'}\nProject: ${n.project || '-'}\n` +
    `Site: ${n.site || '-'}\n\nWrite-up:\n${n.note}`;
  // Opus 5 can decline a request; "default" fallbacks re-run it on Anthropic's recommended
  // substitute server-side instead of handing back a refusal. Only offered on that family.
  const fallback = /^claude-(opus-5|fable)/.test(MODEL);
  try {
    const res = await c.beta.messages.create({
      model: MODEL,
      max_tokens: 16000,
      system: SYSTEM,
      messages: [{ role: 'user', content: prompt }],
      output_config: { effort: 'low', format: { type: 'json_schema', schema: SCHEMA } },
      ...(fallback ? { betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default' } : {}),
    });
    if (res.stop_reason === 'refusal') return { error: 'declined' };
    const text = res.content.filter(b => b.type === 'text').map(b => b.text).join('');
    return { ...JSON.parse(text), model: res.model };
  } catch (e) {
    // Out of credit, bad key, rate limit: keep the note, try the summary again next run.
    return { error: `${e.status ?? ''} ${e.message ?? e}`.trim().slice(0, 200) };
  }
}

/* ---------- the run ---------- */

export async function processWriteups(graph, { dryRun = false } = {}) {
  let notes;
  try { notes = await graph.items(NOTES_LIST); }
  catch (e) {
    console.log(`\nWrite-ups skipped -- is the ${NOTES_LIST} list there? (${e.message.slice(0, 120)})`);
    return;
  }

  const visits = await graph.items(VISIT_LIST).catch(() => []);
  const byNumber = new Map();          // "task|4836" -> visit rows
  for (const it of visits) {
    const f = it.fields ?? {};
    const k = `${f.Kind}|${f.Number}`;
    if (!byNumber.has(k)) byNumber.set(k, []);
    byNumber.get(k).push({ itemId: it.id, ...f });
  }

  let parsed = 0, summarised = 0, matched = 0, skipped = 0, failed = 0;
  for (const it of notes) {
    const f = it.fields ?? {};
    let result = {};
    try { result = JSON.parse(f.Result || '{}'); } catch { result = {}; }
    // Done on an earlier run. A note saved while there was no API key is picked up again
    // once one is added, so it still gets its summary.
    if (f.ProcessedAt && (result.summary || result.skipped || !process.env.ANTHROPIC_API_KEY)) continue;

    const n = parseNotice({ subject: f.Title, body: f.Body, receivedAt: f.ReceivedAt });
    if (!n) { skipped++; if (!dryRun) await graph.update(NOTES_LIST, it.id, { ProcessedAt: new Date().toISOString(), Result: JSON.stringify({ skipped: 'not a site note' }) }).catch(() => {}); continue; }
    parsed++;

    // Summary once; a failure is recorded and retried next run.
    if (!result.summary) {
      const s = await summarise(n);
      if (s.error) failed++; else summarised++;
      result = { ...result, ...s };
    }

    // The visit it belongs to: same kind and number, nearest scheduled date to the note's.
    const candidates = byNumber.get(`${n.kind}|${n.number}`) ?? [];
    const visit = candidates.sort((a, b) =>
      Math.abs(dayGap(a.SchedDate, n.noteDate)) - Math.abs(dayGap(b.SchedDate, n.noteDate)))[0];

    const record = {
      kind: n.kind, number: n.number, by: n.by, client: n.client, project: n.project, site: n.site,
      noteDate: n.noteDate, hours: n.hours, writtenAt: n.writtenAt,
      visitDate: visit?.SchedDate || n.noteDate || '',
      daysLate: daysLate(visit?.SchedDate || n.noteDate, n.writtenAt),
      visitId: visit?.VisitId || '',
    };
    result = { ...result, ...record };

    if (!dryRun) {
      await graph.update(NOTES_LIST, it.id, {
        Result: JSON.stringify(result),
        ...(result.summary || result.error === 'no API key' ? { ProcessedAt: new Date().toISOString() } : {}),
      });
      if (visit) {
        // Several notes can land on one visit (a note, then an addition). Keep them all.
        const prevText = visit.WriteUpText || '';
        const noteBlock = `${n.by} · ${n.writtenAt ? n.writtenAt.slice(0, 10) : ''}\n${n.note}`;
        const text = prevText.includes(n.note) ? prevText : (prevText ? prevText + '\n\n' + noteBlock : noteBlock);
        let summaries = [];
        try { summaries = JSON.parse(visit.Summary || '[]'); if (!Array.isArray(summaries)) summaries = [summaries]; } catch { summaries = []; }
        if (result.summary && !summaries.some(s => s.summary === result.summary)) summaries.push(pick(result));
        await graph.update(VISIT_LIST, visit.itemId, {
          WriteUpAt: earliest(visit.WriteUpAt, n.writtenAt),
          WriteUpBy: visit.WriteUpBy && !visit.WriteUpBy.includes(n.by) ? `${visit.WriteUpBy}, ${n.by}` : (visit.WriteUpBy || n.by),
          WriteUpText: text,
          Summary: JSON.stringify(summaries),
        });
        visit.WriteUpText = text; visit.Summary = JSON.stringify(summaries);
      }
    }
    if (visit) matched++;
  }

  if (parsed || skipped) {
    console.log(`\nwrite-ups: ${parsed} read  ${summarised} summarised  ${matched} matched to a visit` +
      (failed ? `  ${failed} summaries failed (retried next run)` : '') + (skipped ? `  ${skipped} not site notes` : ''));
  }
  if (failed && !(await claude())) console.log('  (no ANTHROPIC_API_KEY secret -- notes are saved and matched, just not summarised)');
}

function pick(r) {
  return { summary: r.summary, completed: r.completed, outstanding: r.outstanding, returnVisit: r.returnVisit,
    officeActions: r.officeActions, clientBuilderNotes: r.clientBuilderNotes, by: r.by, writtenAt: r.writtenAt, daysLate: r.daysLate };
}
function earliest(a, b) { if (!a) return b || ''; if (!b) return a; return a < b ? a : b; }
function dayGap(a, b) {
  if (!a || !b) return 9999;
  return Math.round((new Date(a + 'T00:00:00') - new Date(b + 'T00:00:00')) / 86400000);
}
