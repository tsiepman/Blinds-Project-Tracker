// Work-order write-ups — the "Site Notes added in …" emails from D-Tools Mobile Install.
//
// An installer writing a site note on their phone makes D-Tools send a notification. The
// text of that note is the only place the write-up exists: the API carries Terry's
// instructions for a visit, never what the installer reported afterwards.
//
// Every notice has the same shape:
//
//   Subject: Site Notes added in Task 4836            <- or "Service Order 6306"
//   A Site Note has been added to Task 4836 <link> in Mobile Install by Andrew Lindquist.
//   Client: Bergen, Shawn
//   Project: Wellington Crescent Renovation
//   Name: 1063 Wellington Cres.
//   ==================
//   AL 9/11/2026 (4 hrs).
//   30 min travel
//   Retro 1 cat 6 from Todd's bedroom desk ...
//   ==================
//
// The number in the subject is the task or service-order number, which is what the field
// visit rows are keyed on — so a note matches its visit exactly, with no name guessing.
//
// Notes often arrive days after the visit (one sample: a 30 June visit written up on
// 20 July). The gap is worth keeping: it is the habit problem, measured.

/** Parse one notification email. Returns null if it isn't one. */
export function parseNotice({ subject = '', body = '', receivedAt = '' } = {}) {
  const text = String(body).replace(/\r/g, '');
  const subj = String(subject).replace(/^\s*(FW|RE|FWD):\s*/i, '').trim();

  // The kind and number come from the subject, and are confirmed by the body line.
  const m = subj.match(/Site Notes? added in (Task|Service Order)\s+(\d+)/i)
    || text.match(/has been added to (Task|Service Order)\s+(\d+)/i);
  if (!m) return null;

  const line = (label) => {
    // Stop at the newline: an empty "Project:" must not swallow the "Name:" line below it.
    const r = new RegExp('^' + label + ':[ \\t]*([^\\n]*)$', 'mi').exec(text);
    return r ? r[1].trim() : '';
  };
  const by = (/in Mobile Install by ([^.\n]+)\./i.exec(text) || [, ''])[1].trim();

  // The note itself sits between rows of = signs. Take every block, in case a visit got
  // more than one note in the same email.
  const blocks = [...text.matchAll(/={6,}\n([\s\S]*?)\n={6,}/g)].map(x => x[1].trim()).filter(Boolean);

  // A forwarded copy (Terry's mailbox rule) keeps D-Tools' original send time in the
  // quoted header -- that is when the installer actually wrote the note, so it wins over
  // the forward's own arrival time. A note sent straight to the mailbox has no such line.
  // Forwarded twice (someone forwarding Terry's copy) there are several; the original is
  // the last one before the notice itself.
  const at = text.search(/has been added to (Task|Service Order)/i);
  const sents = [...text.slice(0, at < 0 ? undefined : at).matchAll(/^Sent:[ \t]*([^\n]+)$/gmi)];
  const fwd = sents[sents.length - 1];
  const sentAt = fwd && !isNaN(Date.parse(fwd[1].trim())) ? new Date(fwd[1].trim()).toISOString() : '';
  const arrived = sentAt || receivedAt;

  return {
    kind: /task/i.test(m[1]) ? 'task' : 'service',
    number: m[2],
    by,
    client: line('Client'),
    project: line('Project'),
    site: line('Name'),
    note: blocks.join('\n\n'),
    noteDate: noteDate(blocks[0] || '', arrived),
    hours: noteHours(blocks[0] || ''),
    receivedAt: receivedAt || '',
    sentAt,
    writtenAt: arrived,     // best estimate of when the installer wrote it
  };
}

/* Installers head their notes with their initials, the date they were on site, and the
   hours: "CT - 9/16/26 - 1hr", "BK 7/29/26 5.5 hrs", "YR 9/8/2026 (7 hrs)", "km-06/29-8.75hr".
   The date says which visit the note is really about, which can be weeks before the email. */
export function noteDate(note, arrivedAt) {
  const head = String(note).slice(0, 120);
  const full = /(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/.exec(head);
  // "KM-06/30-8HR" -- month and day only. The year comes from when the note arrived,
  // stepping back a year if that would put the visit in the future.
  const short = !full && /(?:^|[^\d/])(\d{1,2})[/-](\d{1,2})(?![\d/])/.exec(head);
  const m = full || short;
  if (!m) return '';
  const month = Number(m[1]), day = Number(m[2]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return '';
  let year;
  if (full) year = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
  else {
    const ref = arrivedAt && !isNaN(Date.parse(arrivedAt)) ? new Date(arrivedAt) : new Date();
    year = ref.getFullYear();
    if (new Date(year, month - 1, day) > ref) year--;
  }
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

/** Hours from the note heading, when they wrote them. */
export function noteHours(note) {
  const head = String(note).slice(0, 160);
  const h = /(\d+(?:\.\d+)?)\s*(?:hrs?|hours?)\b/i.exec(head);
  if (h) return Number(h[1]);
  const min = /(\d+)\s*(?:min(?:ute)?s?)\b/i.exec(head);     // "CM 9/16/26 30 min"
  return min ? Math.round((Number(min[1]) / 60) * 100) / 100 : null;
}

/** How late the write-up was, in days: visit date to the day the note arrived. */
export function daysLate(visitDate, receivedAt) {
  if (!visitDate || !receivedAt) return null;
  const a = new Date(visitDate + 'T00:00:00');
  const b = new Date(String(receivedAt).slice(0, 10) + 'T00:00:00');
  if (isNaN(a) || isNaN(b)) return null;
  return Math.max(0, Math.round((b - a) / 86400000));
}
