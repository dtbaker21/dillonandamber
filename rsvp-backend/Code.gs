/**
 * Dillon & Amber — RSVP backend
 * Lives inside the RSVP Google Sheet (Extensions > Apps Script) and is deployed as a web app.
 * The wedding website calls it to look up an invitation and to save replies.
 */

const CONFIG = {
  coupleNames: 'Dillon & Amber',
  date: 'Saturday, June 5, 2027',
  venue: 'The Lookout Lodge, 2410 Newberg Rd, Snohomish, WA',
  site: 'https://dtbaker21.github.io/dillonandamber/',
  deadline: new Date('2027-04-16T07:00:00Z'),   // end of April 15, Pacific
  notifyEmail: '',                              // optional: an address to email on every new reply
  meals: ['Herb-roasted chicken', 'Wild mushroom risotto']
};

const GUESTS = 'Guests', LOG = 'Responses', SUMMARY = 'Summary';
const HEAD = ['Invite ID', 'Party name', 'First name', 'Last name', 'Guest type', 'Email',
  'Attending', 'Meal', 'Dietary needs', 'Friday drop-in', 'Song request', 'Note to couple', 'Responded at'];
const C = Object.fromEntries(HEAD.map((h, i) => [h, i]));

/* ---------- Sheet menu ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Wedding RSVP')
    .addItem('Set up or repair this sheet', 'setup')
    .addSeparator()
    .addItem('Email a reminder to everyone who hasn’t replied', 'sendReminders')
    .addToUi();
}

/* ---------- One-time setup: tabs, headers, formatting, sample rows, live summary ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let g = ss.getSheetByName(GUESTS);
  if (!g) {
    g = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(GUESTS) : ss.insertSheet(GUESTS);
  }
  if (g.getLastRow() === 0) {
    g.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
    const sample = [
      ['SAMPLE-1', 'The Baker family', 'Pat', 'Baker', 'Guest', '', '', '', '', '', '', '', ''],
      ['SAMPLE-1', 'The Baker family', 'Chris', 'Baker', 'Guest', '', '', '', '', '', '', '', ''],
      ['SAMPLE-1', 'The Baker family', 'Riley', 'Baker', 'Child', '', '', '', '', '', '', '', ''],
      ['SAMPLE-2', 'Casey Nguyen and guest', 'Casey', 'Nguyen', 'Guest', '', '', '', '', '', '', '', ''],
      ['SAMPLE-2', 'Casey Nguyen and guest', '', '', 'Plus-one', '', '', '', '', '', '', '', ''],
      ['SAMPLE-3', 'Sam Ortiz', 'Sam', 'Ortiz', 'Guest', '', '', '', '', '', '', '', '']
    ];
    g.getRange(2, 1, sample.length, HEAD.length).setValues(sample);
  }
  g.setFrozenRows(1);
  g.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#24382c').setFontColor('#fbf8f3');
  g.setColumnWidths(1, HEAD.length, 140);
  g.setColumnWidth(C['Party name'] + 1, 200);
  g.setColumnWidth(C['Note to couple'] + 1, 260);
  const rules = [
    [C['Guest type'], ['Guest', 'Plus-one', 'Child']],
    [C['Attending'], ['Yes', 'No']],
    [C['Meal'], CONFIG.meals],
    [C['Friday drop-in'], ['Yes', 'No']]
  ];
  rules.forEach(([col, list]) => g.getRange(2, col + 1, 1000, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build()));
  const att = g.getRange(2, C['Attending'] + 1, 1000, 1);
  g.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Yes').setBackground('#dfe9df').setRanges([att]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No').setBackground('#f1e3d9').setRanges([att]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($C2<>"",$M2="")').setBackground('#fbf3e2')
      .setRanges([g.getRange(2, C['Responded at'] + 1, 1000, 1)]).build()
  ]);

  let log = ss.getSheetByName(LOG) || ss.insertSheet(LOG);
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, 6).setValues([['Received', 'Invite ID', 'Party name', 'Who replied', 'Summary', 'Email']]);
    log.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#24382c').setFontColor('#fbf8f3');
    log.setFrozenRows(1); log.setColumnWidth(5, 420);
  }

  let s = ss.getSheetByName(SUMMARY) || ss.insertSheet(SUMMARY, 0);
  s.clear();
  const G = `'${GUESTS}'!`;
  const rows = [
    ['RSVP summary', ''],
    ['Invitations sent', `=COUNTUNIQUE(FILTER(${G}A2:A, ${G}A2:A<>""))`],
    ['Guests invited', `=COUNTA(FILTER(${G}E2:E, ${G}A2:A<>""))`],
    ['Replied', `=COUNTIF(${G}G2:G,"?*")`],
    ['Attending', `=COUNTIF(${G}G2:G,"Yes")`],
    ['Declined', `=COUNTIF(${G}G2:G,"No")`],
    ['Still waiting on', `=B3-B4`],
    ['', ''],
    ['Dinner', ''],
    [CONFIG.meals[0], `=COUNTIFS(${G}G2:G,"Yes",${G}H2:H,"${CONFIG.meals[0]}")`],
    [CONFIG.meals[1], `=COUNTIFS(${G}G2:G,"Yes",${G}H2:H,"${CONFIG.meals[1]}")`],
    ['Kids', `=COUNTIFS(${G}G2:G,"Yes",${G}E2:E,"Child")`],
    ['', ''],
    ['Friday welcome drop-in', `=COUNTIFS(${G}G2:G,"Yes",${G}J2:J,"Yes")`],
    ['Dietary notes', `=COUNTIFS(${G}G2:G,"Yes",${G}I2:I,"?*")`]
  ];
  s.getRange(1, 1, rows.length, 2).setValues(rows);
  s.getRange('A1').setFontSize(16).setFontWeight('bold');
  s.getRange('A9').setFontWeight('bold');
  s.setColumnWidth(1, 240); s.setColumnWidth(2, 120);
  s.getRange('B2:B15').setFontWeight('bold').setHorizontalAlignment('right');
  ss.setActiveSheet(g);
  return 'ok';
}

/* ---------- Web app ---------- */
function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'lookup') return out(lookup(p.first, p.last));
    return out({ ok: true, service: 'rsvp' });
  } catch (err) { return out({ error: 'Something went wrong. Please try again.' }); }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'submit') return out(submit(body));
    return out({ error: 'Unknown request.' });
  } catch (err) { return out({ error: 'We couldn’t save your reply. Please try again.' }); }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z]/g, '');
const matches = (r, f, l) => { const gf = norm(r[C['First name']]), gl = norm(r[C['Last name']]); return gl && gl === l && gf && (gf === f || gf.startsWith(f) || f.startsWith(gf)); };

function table() {
  const g = SpreadsheetApp.getActive().getSheetByName(GUESTS);
  const v = g.getDataRange().getValues();
  return { g, rows: v.slice(1).map((r, i) => ({ n: i + 2, r })).filter(x => x.r[C['Invite ID']] !== '') };
}

function pack(rows, id) {
  // Privacy: a lookup only reveals who is on the invitation, never their answers.
  const mine = rows.filter(x => String(x.r[C['Invite ID']]) === String(id));
  const times = mine.map(x => x.r[C['Responded at']]).filter(t => t instanceof Date);
  const last = times.length ? new Date(Math.max.apply(null, times)) : null;
  return {
    invite: String(id), party: mine[0].r[C['Party name']] || '',
    responded: !!last,
    respondedOn: last ? Utilities.formatDate(last, 'America/Los_Angeles', 'MMMM d') : '',
    guests: mine.map(x => ({ row: x.n, first: x.r[C['Guest type']] === 'Plus-one' ? '' : x.r[C['First name']],
      last: x.r[C['Guest type']] === 'Plus-one' ? '' : x.r[C['Last name']], type: x.r[C['Guest type']] || 'Guest' }))
  };
}

function fullInvite(rows, id) {
  const mine = rows.filter(x => String(x.r[C['Invite ID']]) === String(id));
  return { party: mine[0].r[C['Party name']], guests: mine.map(x => ({ first: x.r[C['First name']], last: x.r[C['Last name']], attending: x.r[C['Attending']], meal: x.r[C['Meal']] })) };
}

/* Gentle brute-force protection: too many misses in a short window pauses lookups. */
function tooManyMisses(add) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('misses') || 0) + (add ? 1 : 0);
  if (add) cache.put('misses', String(n), 600);
  return n >= 40;
}

function lookup(first, last) {
  if (new Date() > CONFIG.deadline) return { closed: true };
  if (tooManyMisses(false)) return { paused: true };
  const f = norm(first), l = norm(last);
  if (!f || l.length < 2) return { found: false };
  const { rows } = table();
  const ids = [...new Set(rows.filter(x => matches(x.r, f, l)).map(x => String(x.r[C['Invite ID']])))];
  if (!ids.length) { tooManyMisses(true); return { found: false }; }
  return { found: true, invites: ids.map(id => pack(rows, id)) };
}

function submit(b) {
  if (new Date() > CONFIG.deadline) return { error: 'RSVPs are closed. Please text or email us.' };
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const { g, rows } = table();
    const mine = rows.filter(x => String(x.r[C['Invite ID']]) === String(b.invite));
    if (!mine.length) return { error: 'We couldn’t find your invitation.' };
    const f = norm(b.verify && b.verify.first), l = norm(b.verify && b.verify.last);
    if (!mine.some(x => matches(x.r, f, l))) return { error: 'Please look yourself up again.' };
    const now = new Date();
    const attending = [];
    (b.guests || []).forEach(gu => {
      const x = mine.find(m => m.n === Number(gu.row)); if (!x) return;
      const yes = gu.attending === 'Yes';
      const meal = yes && CONFIG.meals.includes(gu.meal) ? gu.meal : '';
      const vals = x.r.slice();
      if (vals[C['Guest type']] === 'Plus-one') { vals[C['First name']] = clean(gu.first); vals[C['Last name']] = clean(gu.last); }
      vals[C['Attending']] = yes ? 'Yes' : gu.attending === 'No' ? 'No' : '';
      vals[C['Meal']] = meal;
      vals[C['Dietary needs']] = yes ? clean(gu.diet) : '';
      vals[C['Friday drop-in']] = b.friday === 'Yes' || b.friday === 'No' ? b.friday : '';
      vals[C['Song request']] = clean(b.song);
      vals[C['Note to couple']] = clean(b.note, 1000);
      vals[C['Email']] = clean(b.email) || vals[C['Email']];
      vals[C['Responded at']] = now;
      g.getRange(x.n, 1, 1, HEAD.length).setValues([vals]);
      x.r = vals;
      attending.push(`${vals[C['First name']] || 'Guest'}: ${yes ? meal : 'not attending'}`);
    });
    const party = mine[0].r[C['Party name']];
    SpreadsheetApp.getActive().getSheetByName(LOG)
      .appendRow([now, b.invite, party, `${clean(b.verify.first)} ${clean(b.verify.last)}`, attending.join(' · ') + (b.friday ? ` · Friday: ${b.friday}` : ''), clean(b.email)]);
    SpreadsheetApp.flush();
    const email = clean(b.email);
    try { if (email) confirmEmail(email, fullInvite(table().rows, b.invite)); } catch (e) {}
    try { if (CONFIG.notifyEmail) MailApp.sendEmail(CONFIG.notifyEmail, `New RSVP: ${party}`, attending.join('\n')); } catch (e) {}
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function clean(s, max) { return String(s || '').replace(/[<>]/g, '').trim().slice(0, max || 200); }

/* ---------- Emails ---------- */
function shell(inner) {
  return `<div style="background:#f4eee6;padding:32px 16px;font-family:Georgia,serif;color:#2a211c">
    <div style="max-width:520px;margin:0 auto;background:#fbf8f3;border-radius:18px;padding:32px">
      <p style="margin:0 0 4px;font:600 13px Helvetica,Arial,sans-serif;color:#5f7a5c">${CONFIG.date}</p>
      <h1 style="margin:0 0 18px;font-weight:400;font-size:30px">${CONFIG.coupleNames}</h1>${inner}
      <p style="margin:24px 0 0;font:14px Helvetica,Arial,sans-serif;color:#76695e">${CONFIG.venue}<br><a href="${CONFIG.site}" style="color:#9a6a2e">${CONFIG.site.replace('https://', '')}</a></p>
    </div></div>`;
}

function confirmEmail(to, inv) {
  const yes = inv.guests.filter(x => x.attending === 'Yes');
  const list = inv.guests.map(x => `<tr><td style="padding:8px 0;border-bottom:1px solid #e6dccd">${x.first || 'Guest'} ${x.last || ''}</td><td style="padding:8px 0;border-bottom:1px solid #e6dccd;text-align:right;color:#76695e">${x.attending === 'Yes' ? x.meal : 'Not attending'}</td></tr>`).join('');
  const inner = `<p style="font:16px/1.55 Helvetica,Arial,sans-serif;margin:0 0 16px">${yes.length ? 'You’re all set. We can’t wait to celebrate with you!' : 'Thank you for letting us know. We’ll miss you!'}</p>
    <table style="width:100%;border-collapse:collapse;font:15px Helvetica,Arial,sans-serif">${list}</table>
    <p style="font:14px Helvetica,Arial,sans-serif;color:#76695e;margin:16px 0 0">Need to change something? Look yourself up again on our website any time before April 15.</p>`;
  MailApp.sendEmail({ to, subject: `Your RSVP for ${CONFIG.coupleNames}`, htmlBody: shell(inner), name: CONFIG.coupleNames });
}

function sendReminders() {
  const ui = SpreadsheetApp.getUi();
  const { rows } = table();
  const byInvite = {};
  rows.forEach(x => { const id = String(x.r[C['Invite ID']]); (byInvite[id] = byInvite[id] || []).push(x); });
  const pending = Object.values(byInvite).filter(list => !list.some(x => x.r[C['Responded at']] !== '') && list.some(x => x.r[C['Email']]));
  if (!pending.length) { ui.alert('Everyone with an email address has replied.'); return; }
  const quota = MailApp.getRemainingDailyQuota();
  const ok = ui.alert(`Send a reminder to ${pending.length} invitation(s)?`, `Google lets this account send about ${quota} more emails today.`, ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  let sent = 0;
  pending.slice(0, quota).forEach(list => {
    const to = list.find(x => x.r[C['Email']]).r[C['Email']];
    const inner = `<p style="font:16px/1.55 Helvetica,Arial,sans-serif;margin:0 0 18px">We’re finalizing numbers and would love to know if you can make it. Please RSVP by April 15. It takes about a minute.</p>
      <a href="${CONFIG.site}#rsvp" style="display:inline-block;background:#24382c;color:#fbf8f3;text-decoration:none;font:600 15px Helvetica,Arial,sans-serif;padding:12px 22px;border-radius:999px">RSVP now</a>`;
    MailApp.sendEmail({ to, subject: `A quick RSVP reminder from ${CONFIG.coupleNames}`, htmlBody: shell(inner), name: CONFIG.coupleNames });
    sent++;
  });
  ui.alert(`Sent ${sent} reminder(s).${pending.length > sent ? ` ${pending.length - sent} more will need to go out tomorrow.` : ''}`);
}
