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
  'Attending', 'Meal', 'Dietary needs', 'Friday drop-in', 'Song request', 'Note to couple', 'Responded at',
  'Staying at', 'Shuttle', 'Rehearsal invite', 'Rehearsal dinner', 'Brunch', 'Mailing address', 'Phone',
  'Save the date sent', 'Invitation sent', 'Gift', 'Thank-you sent'];
const L = name => { let n = C[name] + 1, s = ''; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
const SITE = 'Site', ADDR = 'Addresses';
const STAYS = ['Our hotel block', 'Another hotel nearby', 'A rental or B&B', 'With friends or family', 'Heading home that night'];
const SITE_DEFAULTS = [
  ['Setting', 'Value', 'What it does'],
  ['Announcement', '', 'Shows as a banner across the top of the website. Leave blank to hide it.'],
  ['Announcement link', '', 'Optional link for the banner (for example a map or a page).'],
  ['Hotel block name', '', 'Name of the hotel holding your room block.'],
  ['Hotel block link', '', 'The hotel’s group booking link. Guests tap Book a room.'],
  ['Hotel block note', 'Group rate held until early May. It’s also the shuttle stop.', 'One line under the hotel name.'],
  ['Amazon registry', '', 'Link to your Amazon registry.'],
  ['Crate & Barrel registry', '', 'Link to your Crate & Barrel registry.'],
  ['Other registry', '', 'Any other registry link (Zola, Target, etc).'],
  ['Venmo username', '', 'For the honeymoon fund, without the @. No fees for guests paying from a bank or balance.'],
  ['Photo album link', '', 'Optional: a shared link to the guest photos folder, if you want guests to see everyone’s photos.'],
  ['Brunch details', '', 'Sunday brunch: when this is filled in, the website shows it in the schedule and the RSVP asks about it.']
];
const C = Object.fromEntries(HEAD.map((h, i) => [h, i]));

/* ---------- Sheet menu ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Wedding RSVP')
    .addItem('Set up or repair this sheet', 'setup')
    .addSeparator()
    .addItem('Email a reminder to everyone who hasn’t replied', 'sendReminders')
    .addItem('Email the save-the-date', 'sendSaveTheDates')
    .addItem('Email the invitation', 'sendInvitations')
    .addSeparator()
    .addItem('Copy new addresses into the guest list', 'mergeAddresses')
    .addToUi();
}

/* ---------- Setup and repair: tabs, headers, formatting, sample rows, live summary ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let g = ss.getSheetByName(GUESTS);
  if (!g) g = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(GUESTS) : ss.insertSheet(GUESTS);
  if (g.getLastRow() === 0) {
    const row = (id, party, first, last, type, extra) => { const r = HEAD.map(() => ''); Object.assign(r, { 0: id, 1: party, 2: first, 3: last, 4: type }); Object.entries(extra || {}).forEach(([k, v]) => r[C[k]] = v); return r; };
    const sample = [
      row('SAMPLE-1', 'The Baker family', 'Pat', 'Baker', 'Guest', { 'Rehearsal invite': 'Yes' }),
      row('SAMPLE-1', 'The Baker family', 'Chris', 'Baker', 'Guest', { 'Rehearsal invite': 'Yes' }),
      row('SAMPLE-1', 'The Baker family', 'Riley', 'Baker', 'Child', { 'Rehearsal invite': 'Yes' }),
      row('SAMPLE-2', 'Casey Nguyen and guest', 'Casey', 'Nguyen', 'Guest'),
      row('SAMPLE-2', 'Casey Nguyen and guest', '', '', 'Plus-one'),
      row('SAMPLE-3', 'Sam Ortiz', 'Sam', 'Ortiz', 'Guest')
    ];
    g.getRange(2, 1, sample.length, HEAD.length).setValues(sample);
  }
  if (g.getMaxColumns() < HEAD.length) g.insertColumnsAfter(g.getMaxColumns(), HEAD.length - g.getMaxColumns());
  g.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#24382c').setFontColor('#fbf8f3');
  g.setFrozenRows(1);
  g.setColumnWidths(1, HEAD.length, 140);
  g.setColumnWidth(C['Party name'] + 1, 200);
  g.setColumnWidth(C['Note to couple'] + 1, 260);
  g.setColumnWidth(C['Mailing address'] + 1, 280);
  const yn = ['Yes', 'No'];
  [[C['Guest type'], ['Guest', 'Plus-one', 'Child']], [C['Attending'], yn], [C['Meal'], CONFIG.meals], [C['Friday drop-in'], yn],
   [C['Staying at'], STAYS], [C['Shuttle'], yn], [C['Rehearsal invite'], yn], [C['Rehearsal dinner'], yn], [C['Brunch'], yn], [C['Thank-you sent'], yn]]
    .forEach(([col, list]) => g.getRange(2, col + 1, 1000, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build()));
  const att = g.getRange(2, C['Attending'] + 1, 1000, 1);
  g.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Yes').setBackground('#dfe9df').setRanges([att]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No').setBackground('#f1e3d9').setRanges([att]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($C2<>"",$${L('Responded at')}2="")`).setBackground('#fbf3e2')
      .setRanges([g.getRange(2, C['Responded at'] + 1, 1000, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($${L('Gift')}2<>"",$${L('Thank-you sent')}2<>"Yes")`).setBackground('#fbf3e2')
      .setRanges([g.getRange(2, C['Thank-you sent'] + 1, 1000, 1)]).build()
  ]);

  const tab = (name, head, width) => {
    const t = ss.getSheetByName(name) || ss.insertSheet(name);
    if (t.getLastRow() === 0) {
      t.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#24382c').setFontColor('#fbf8f3');
      t.setFrozenRows(1); (width || []).forEach(([c, w]) => t.setColumnWidth(c, w));
    }
    return t;
  };
  tab(LOG, ['Received', 'Invite ID', 'Party name', 'Who replied', 'Summary', 'Email'], [[5, 420]]);
  tab(ADDR, ['Received', 'First name', 'Last name', 'Household', 'Street', 'Apt', 'City', 'State', 'ZIP', 'Country', 'Email', 'Phone', 'Matched invite', 'Copied to guest list'], [[4, 200], [5, 220]]);

  const site = ss.getSheetByName(SITE) || ss.insertSheet(SITE);
  if (site.getLastRow() === 0) {
    site.getRange(1, 1, SITE_DEFAULTS.length, 3).setValues(SITE_DEFAULTS);
  } else {
    const have = site.getRange(1, 1, site.getLastRow(), 1).getValues().map(r => r[0]);
    SITE_DEFAULTS.filter(r => !have.includes(r[0])).forEach(r => site.appendRow(r));
  }
  site.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#24382c').setFontColor('#fbf8f3');
  site.setFrozenRows(1); site.setColumnWidth(1, 200); site.setColumnWidth(2, 360); site.setColumnWidth(3, 460);
  site.getRange(2, 3, site.getLastRow(), 1).setFontColor('#76695e').setWrap(true);

  const folder = photoFolder();

  let s = ss.getSheetByName(SUMMARY) || ss.insertSheet(SUMMARY, 0);
  s.clear();
  const G = `'${GUESTS}'!`, col = n => `${G}${L(n)}2:${L(n)}`;
  const yes = (n, v) => `=COUNTIFS(${col('Attending')},"Yes",${col(n)},"${v || 'Yes'}")`;
  const rows = [
    ['Wedding summary', ''],
    ['Invitations', `=COUNTUNIQUE(FILTER(${col('Invite ID')}, ${col('Invite ID')}<>""))`],
    ['Guests invited', `=COUNTA(FILTER(${col('Guest type')}, ${col('Invite ID')}<>""))`],
    ['Replied', `=COUNTIF(${col('Attending')},"?*")`],
    ['Attending', `=COUNTIF(${col('Attending')},"Yes")`],
    ['Declined', `=COUNTIF(${col('Attending')},"No")`],
    ['Still waiting on', `=B3-B4`],
    ['', ''],
    ['Dinner', ''],
    [CONFIG.meals[0], yes('Meal', CONFIG.meals[0])],
    [CONFIG.meals[1], yes('Meal', CONFIG.meals[1])],
    ['Kids', yes('Guest type', 'Child')],
    ['Dietary notes', `=COUNTIFS(${col('Attending')},"Yes",${col('Dietary needs')},"?*")`],
    ['', ''],
    ['Events', ''],
    ['Friday drop-in', yes('Friday drop-in')],
    ['Rehearsal dinner', yes('Rehearsal dinner')],
    ['Sunday brunch', yes('Brunch')],
    ['Riding the shuttle', yes('Shuttle')],
    ['', ''],
    ['Where guests are staying', ''],
    ...STAYS.map(x => [x, yes('Staying at', x)]),
    ['', ''],
    ['Mail and gifts', ''],
    ['Addresses collected', `=COUNTA(FILTER(${col('Mailing address')}, ${col('Mailing address')}<>""))`],
    ['Save-the-dates emailed', `=COUNTUNIQUE(FILTER(${col('Invite ID')}, ${col('Save the date sent')}<>""))`],
    ['Invitations emailed', `=COUNTUNIQUE(FILTER(${col('Invite ID')}, ${col('Invitation sent')}<>""))`],
    ['Gifts logged', `=COUNTIF(${col('Gift')},"?*")`],
    ['Thank-yous still to write', `=COUNTIFS(${col('Gift')},"?*",${col('Thank-you sent')},"<>Yes")`],
    ['', ''],
    ['Guest photos folder', folder.getUrl()]
  ];
  s.getRange(1, 1, rows.length, 2).setValues(rows);
  s.getRange('A1').setFontSize(16).setFontWeight('bold');
  rows.forEach((r, i) => { if (r[0] && r[1] === '' && i) s.getRange(i + 1, 1).setFontWeight('bold'); });
  s.setColumnWidth(1, 240); s.setColumnWidth(2, 120);
  s.getRange(2, 2, rows.length - 2, 1).setFontWeight('bold').setHorizontalAlignment('right');
  ss.setActiveSheet(g);
  return 'ok';
}

function photoFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('photoFolder');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(`${CONFIG.coupleNames} · guest photos`);
  props.setProperty('photoFolder', f.getId());
  return f;
}

/* ---------- Web app ---------- */
function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'lookup') return out(lookup(p.first, p.last));
    if (p.action === 'site') return out(siteInfo());
    return out({ ok: true, service: 'rsvp' });
  } catch (err) { return out({ error: 'Something went wrong. Please try again.' }); }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'submit') return out(submit(body));
    if (body.action === 'address') return out(saveAddress(body));
    if (body.action === 'photo') return out(savePhoto(body));
    return out({ error: 'Unknown request.' });
  } catch (err) { return out({ error: 'We couldn’t save that. Please try again.' }); }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z]/g, '');
const matches = (r, f, l) => { const gf = norm(r[C['First name']]), gl = norm(r[C['Last name']]); return gl && gl === l && gf && (gf === f || gf.startsWith(f) || f.startsWith(gf)); };

function table() {
  const g = SpreadsheetApp.getActive().getSheetByName(GUESTS);
  const v = g.getRange(1, 1, Math.max(g.getLastRow(), 1), HEAD.length).getValues();
  return { g, rows: v.slice(1).map((r, i) => ({ n: i + 2, r })).filter(x => x.r[C['Invite ID']] !== '') };
}

function settings() {
  const t = SpreadsheetApp.getActive().getSheetByName(SITE);
  if (!t || t.getLastRow() < 2) return {};
  return Object.fromEntries(t.getRange(2, 1, t.getLastRow() - 1, 2).getValues().filter(r => r[0]).map(r => [r[0], String(r[1]).trim()]));
}

/* Public site settings, plus where attending guests are staying (counts only, no names). */
function siteInfo() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('site'); if (hit) return JSON.parse(hit);
  const s = settings();
  const staying = {};
  table().rows.forEach(x => { if (x.r[C['Attending']] === 'Yes' && x.r[C['Staying at']]) staying[x.r[C['Staying at']]] = (staying[x.r[C['Staying at']]] || 0) + 1; });
  const res = {
    announcement: s['Announcement'] || '', announcementLink: s['Announcement link'] || '',
    hotel: { name: s['Hotel block name'] || '', link: s['Hotel block link'] || '', note: s['Hotel block note'] || '' },
    registry: { amazon: s['Amazon registry'] || '', crate: s['Crate & Barrel registry'] || '', other: s['Other registry'] || '', venmo: (s['Venmo username'] || '').replace(/^@/, '') },
    album: s['Photo album link'] || '', brunch: s['Brunch details'] || '', staying
  };
  cache.put('site', JSON.stringify(res), 60);
  return res;
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
    rehearsal: mine.some(x => x.r[C['Rehearsal invite']] === 'Yes'),
    brunch: !!settings()['Brunch details'],
    guests: mine.map(x => ({ row: x.n, first: x.r[C['Guest type']] === 'Plus-one' ? '' : x.r[C['First name']],
      last: x.r[C['Guest type']] === 'Plus-one' ? '' : x.r[C['Last name']], type: x.r[C['Guest type']] || 'Guest' }))
  };
}

function fullInvite(rows, id) {
  const mine = rows.filter(x => String(x.r[C['Invite ID']]) === String(id));
  return { party: mine[0].r[C['Party name']], guests: mine.map(x => ({ first: x.r[C['First name']], last: x.r[C['Last name']], attending: x.r[C['Attending']], meal: x.r[C['Meal']] })) };
}

/* Gentle abuse protection: counters that pause a feature after too many hits in a short window. */
function overLimit(key, add, max, secs) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(key) || 0) + (add ? 1 : 0);
  if (add) cache.put(key, String(n), secs);
  return n >= max;
}

function lookup(first, last) {
  if (new Date() > CONFIG.deadline) return { closed: true };
  if (overLimit('misses', false, 40, 600)) return { paused: true };
  const f = norm(first), l = norm(last);
  if (!f || l.length < 2) return { found: false };
  const { rows } = table();
  const ids = [...new Set(rows.filter(x => matches(x.r, f, l)).map(x => String(x.r[C['Invite ID']])))];
  if (!ids.length) { overLimit('misses', true, 40, 600); return { found: false }; }
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
    const invited = mine.some(x => x.r[C['Rehearsal invite']] === 'Yes');
    const yn = v => v === 'Yes' || v === 'No' ? v : '';
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
      vals[C['Friday drop-in']] = yn(b.friday);
      vals[C['Staying at']] = STAYS.includes(b.staying) ? b.staying : '';
      vals[C['Shuttle']] = yn(b.shuttle);
      vals[C['Rehearsal dinner']] = invited ? yn(b.rehearsal) : '';
      vals[C['Brunch']] = yn(b.brunch);
      vals[C['Song request']] = clean(b.song);
      vals[C['Note to couple']] = clean(b.note, 1000);
      vals[C['Email']] = clean(b.email) || vals[C['Email']];
      vals[C['Responded at']] = now;
      g.getRange(x.n, 1, 1, HEAD.length).setValues([vals]);
      x.r = vals;
      attending.push(`${vals[C['First name']] || 'Guest'}: ${yes ? meal : 'not attending'}`);
    });
    const party = mine[0].r[C['Party name']];
    const extras = [b.friday && `Friday: ${b.friday}`, b.staying && `Staying: ${b.staying}`, b.shuttle && `Shuttle: ${b.shuttle}`,
      invited && b.rehearsal && `Rehearsal: ${b.rehearsal}`, b.brunch && `Brunch: ${b.brunch}`].filter(Boolean);
    SpreadsheetApp.getActive().getSheetByName(LOG)
      .appendRow([now, b.invite, party, `${clean(b.verify.first)} ${clean(b.verify.last)}`, attending.concat(extras).join(' · '), clean(b.email)]);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove('site');
    const email = clean(b.email);
    try { if (email) confirmEmail(email, fullInvite(table().rows, b.invite)); } catch (e) {}
    try { if (CONFIG.notifyEmail) MailApp.sendEmail(CONFIG.notifyEmail, `New RSVP: ${party}`, attending.concat(extras).join('\n')); } catch (e) {}
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function clean(s, max) { return String(s || '').replace(/[<>]/g, '').trim().slice(0, max || 200); }

/* ---------- Address collection ---------- */
function formatAddress(a) {
  return [a.street + (a.apt ? `, ${a.apt}` : ''), `${a.city}, ${a.state} ${a.zip}`.trim(), a.country && !/^(us|usa|united states)$/i.test(a.country) ? a.country : '']
    .filter(Boolean).join('\n');
}

function copyAddressToGuests(rows, g, a) {
  const f = norm(a.first), l = norm(a.last);
  const ids = [...new Set(rows.filter(x => matches(x.r, f, l)).map(x => String(x.r[C['Invite ID']])))];
  if (ids.length !== 1) return '';
  rows.filter(x => String(x.r[C['Invite ID']]) === ids[0]).forEach(x => {
    g.getRange(x.n, C['Mailing address'] + 1).setValue(formatAddress(a));
    if (a.phone && !x.r[C['Phone']]) g.getRange(x.n, C['Phone'] + 1).setValue(a.phone);
    if (a.email && !x.r[C['Email']] && matches(x.r, f, l)) g.getRange(x.n, C['Email'] + 1).setValue(a.email);
  });
  return ids[0];
}

function saveAddress(b) {
  if (b.website) return { ok: true };                     // bot trap
  if (overLimit('addr', true, 300, 3600)) return { error: 'Too many submissions right now. Please try again later.' };
  const a = {};
  ['first', 'last', 'household', 'street', 'apt', 'city', 'state', 'zip', 'country', 'email', 'phone'].forEach(k => a[k] = clean(b[k]));
  if (!a.first || !a.last || !a.street || !a.city || !a.zip) return { error: 'Please fill in your name and full address.' };
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const { g, rows } = table();
    const id = copyAddressToGuests(rows, g, a);
    SpreadsheetApp.getActive().getSheetByName(ADDR).appendRow([new Date(), a.first, a.last, a.household, a.street, a.apt, a.city, a.state, a.zip, a.country, a.email, a.phone, id, id ? 'Yes' : '']);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function mergeAddresses() {
  const ui = SpreadsheetApp.getUi();
  const t = SpreadsheetApp.getActive().getSheetByName(ADDR);
  if (t.getLastRow() < 2) { ui.alert('No addresses yet.'); return; }
  const { g, rows } = table();
  const data = t.getRange(2, 1, t.getLastRow() - 1, 14).getValues();
  let done = 0, left = 0;
  data.forEach((r, i) => {
    if (r[13] === 'Yes') return;
    const a = { first: r[1], last: r[2], street: r[4], apt: r[5], city: r[6], state: r[7], zip: r[8], country: r[9], email: r[10], phone: r[11] };
    const id = copyAddressToGuests(rows, g, a);
    if (id) { t.getRange(i + 2, 13, 1, 2).setValues([[id, 'Yes']]); done++; } else left++;
  });
  ui.alert(`Copied ${done} address(es) into the guest list.${left ? ` ${left} didn’t match a guest name exactly; add those people to the Guests tab (or fix the spelling) and run this again.` : ''}`);
}

/* ---------- Guest photo uploads ---------- */
function savePhoto(b) {
  if (overLimit('photos', true, 600, 3600)) return { error: 'Lots of uploads right now. Please try again in a bit.' };
  if (!b.data || !/^image\//.test(b.type || '')) return { error: 'Please choose a photo.' };
  const bytes = Utilities.base64Decode(b.data);
  if (bytes.length > 15 * 1024 * 1024) return { error: 'That photo is too large.' };
  const from = clean(b.from) || 'A guest';
  const ext = (b.type.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  const name = `${Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd HH.mm.ss')} · ${from}.${ext}`;
  photoFolder().createFile(Utilities.newBlob(bytes, b.type, name)).setDescription(`From ${from}`);
  return { ok: true };
}

/* ---------- Emails ---------- */
function shell(inner) {
  return `<div style="background:#f4eee6;padding:32px 16px;font-family:Georgia,serif;color:#2a211c">
    <div style="max-width:520px;margin:0 auto;background:#fbf8f3;border-radius:18px;padding:32px">
      <p style="margin:0 0 4px;font:600 13px Helvetica,Arial,sans-serif;color:#5f7a5c">${CONFIG.date}</p>
      <h1 style="margin:0 0 18px;font-weight:400;font-size:30px">${CONFIG.coupleNames}</h1>${inner}
      <p style="margin:24px 0 0;font:14px Helvetica,Arial,sans-serif;color:#76695e">${CONFIG.venue}<br><a href="${CONFIG.site}" style="color:#9a6a2e">${CONFIG.site.replace('https://', '')}</a></p>
    </div></div>`;
}
const button = (href, label) => `<a href="${href}" style="display:inline-block;background:#24382c;color:#fbf8f3;text-decoration:none;font:600 15px Helvetica,Arial,sans-serif;padding:12px 22px;border-radius:999px">${label}</a>`;
const para = t => `<p style="font:16px/1.55 Helvetica,Arial,sans-serif;margin:0 0 18px">${t}</p>`;

function confirmEmail(to, inv) {
  const yes = inv.guests.filter(x => x.attending === 'Yes');
  const list = inv.guests.map(x => `<tr><td style="padding:8px 0;border-bottom:1px solid #e6dccd">${x.first || 'Guest'} ${x.last || ''}</td><td style="padding:8px 0;border-bottom:1px solid #e6dccd;text-align:right;color:#76695e">${x.attending === 'Yes' ? x.meal : 'Not attending'}</td></tr>`).join('');
  const inner = para(yes.length ? 'You’re all set. We can’t wait to celebrate with you!' : 'Thank you for letting us know. We’ll miss you!') +
    `<table style="width:100%;border-collapse:collapse;font:15px Helvetica,Arial,sans-serif">${list}</table>
    <p style="font:14px Helvetica,Arial,sans-serif;color:#76695e;margin:16px 0 0">Need to change something? Look yourself up again on our website any time before April 15.</p>`;
  MailApp.sendEmail({ to, subject: `Your RSVP for ${CONFIG.coupleNames}`, htmlBody: shell(inner), name: CONFIG.coupleNames });
}

/* Send one email per invitation, to the first guest on it with an email address. */
function emailInvites(opts) {
  const ui = SpreadsheetApp.getUi();
  const { g, rows } = table();
  const byInvite = {};
  rows.forEach(x => { const id = String(x.r[C['Invite ID']]); (byInvite[id] = byInvite[id] || []).push(x); });
  const pending = Object.values(byInvite).filter(list => opts.filter(list) && list.some(x => x.r[C['Email']]));
  if (!pending.length) { ui.alert(opts.none); return; }
  const quota = MailApp.getRemainingDailyQuota();
  const ok = ui.alert(`${opts.title} to ${pending.length} invitation(s)?`, `Google lets this account send about ${quota} more emails today.`, ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  let sent = 0;
  pending.slice(0, quota).forEach(list => {
    const to = list.find(x => x.r[C['Email']]).r[C['Email']];
    MailApp.sendEmail({ to, subject: opts.subject, htmlBody: shell(opts.body(list)), name: CONFIG.coupleNames });
    if (opts.mark) list.forEach(x => g.getRange(x.n, C[opts.mark] + 1).setValue(new Date()));
    sent++;
  });
  ui.alert(`Sent ${sent}.${pending.length > sent ? ` ${pending.length - sent} more will need to go out tomorrow (Google’s daily limit).` : ''}`);
}

function sendSaveTheDates() {
  emailInvites({
    title: 'Email the save-the-date', none: 'Everyone with an email address already has one.',
    subject: `Save the date: ${CONFIG.coupleNames}, June 5, 2027`, mark: 'Save the date sent',
    filter: list => !list.some(x => x.r[C['Save the date sent']] !== ''),
    body: () => para('We’re getting married! Please save the date: <b>Saturday, June 5, 2027</b>, at The Lookout Lodge in Snohomish, Washington. A formal invitation will follow.') +
      para('Travel tips, places to stay and our hotel block are on our website.') + button(CONFIG.site, 'Visit our website')
  });
}

function sendInvitations() {
  emailInvites({
    title: 'Email the invitation', none: 'Everyone with an email address already has one.',
    subject: `You’re invited: ${CONFIG.coupleNames}, June 5, 2027`, mark: 'Invitation sent',
    filter: list => !list.some(x => x.r[C['Invitation sent']] !== ''),
    body: list => para(`Dear ${list[0].r[C['Party name']] || 'friends'},`) +
      para('We would love for you to join us as we get married on <b>Saturday, June 5, 2027</b>. The ceremony begins at 4:00 pm at The Lookout Lodge in Snohomish, Washington, with dinner and dancing to follow.') +
      para('Please RSVP by April 15. It takes about a minute.') + button(CONFIG.site + '#rsvp', 'RSVP')
  });
}

function sendReminders() {
  emailInvites({
    title: 'Send a reminder', none: 'Everyone with an email address has replied.',
    subject: `A quick RSVP reminder from ${CONFIG.coupleNames}`,
    filter: list => !list.some(x => x.r[C['Responded at']] !== ''),
    body: () => para('We’re finalizing numbers and would love to know if you can make it. Please RSVP by April 15. It takes about a minute.') + button(CONFIG.site + '#rsvp', 'RSVP now')
  });
}
