/**
 * Saski’O Discover — enquiry backend (Google Apps Script, bound to the Enquiry Register sheet)
 * Owner: hello.saskio@gmail.com · Built by Brad · v1.0 · 18 Nov 2026
 *
 * What it does, in this order (CF04: nothing is confirmed unless saved):
 *   1. Receives an enquiry from the Discover site (POST, JSON as text/plain).
 *   2. Checks it (required fields, lengths, spam trap) and ignores duplicate re-sends.
 *   3. Saves one row to the "Register" tab and flushes.
 *   4. Replies { ok:true, id } — the site only shows "Thank you" after this.
 *   5. Sends the internal notification and the customer acknowledgement; records Sent/Failed.
 *
 * One-off setup: run setup() once from the editor, then Deploy > New deployment > Web app
 * (Execute as: Me · Who has access: Anyone). Paste the web app URL into the site config.
 */

const CONFIG = {
  NOTIFY_TO: 'hello.saskio@gmail.com',
  SHEET_ID: '1JcFkZoNvXB7IrcXS6UKz4Lgxq2YxfUP-GeELsalcojQ',   // Saski’O Discover — Enquiry Register
  SITE_URL: 'https://hellosaskio.github.io/discover/',                         // filled in once the test site address exists
  PRIVACY_VERSION: 'PN v1.0 (18 Nov 2026)',
  TIMEZONE: 'Europe/London',
  SHEET: 'Register', LISTS: 'Lists', DICT: 'Data dictionary', LOG: 'Log'
};

// Columns written by the automation — always first, always in this order.
const AUTO = [
  'Enquiry ID','Submitted (UK)','Timezone','Customer name','Email','Phone','Preferred contact',
  'Destination / ideas','Date preference','Travel from','Travel to','Flexible timing','Duration',
  'Adults (18+)','Children (17 and under)','Child ages at travel','Total travellers','Budget band (total)',
  'Already looking?','Found a place — details','Interests','Additional details / occasion',
  'Source channel','Link / placement','Campaign','Content','Entry route','Click ID present',
  'Privacy notice version shown','Marketing consent','Marketing consent at',
  'Internal notification','Acknowledgement email','Client submission ID'
];
// Columns for Victoria (manual) and calculated columns — after the automated block, never written by the script.
const MANUAL = [
  ['Status','manual'],['Owner','manual'],['Next action','manual'],['Follow-up date','manual'],['Last contacted','manual'],
  ['Est. booking value (£)','manual'],['Gross commission %','manual'],['Saski’O share %','manual'],
  ['Net commission (£)','formula'],['Projected hours','manual'],['Net £ per hour','formula'],['vs £25/hr benchmark','formula'],
  ['Asset-library match','manual'],['Outcome','manual'],['Delete after','formula']
];

const LIMITS = {dest:500, notes:500, found:500, flex:200, name:100, phone:30, email:254};
const LISTS = {
  dateMode:['I have dates in mind','I’m flexible'],
  duration:['','Up to 1 week','1–2 weeks','2–3 weeks','3+ weeks','Not sure yet'],
  budget:['Under £1,000','£1,000 – £2,500','£2,500 – £5,000','£5,000 – £10,000','£10,000+','Not sure yet'],
  looking:['Yes, I’ve found something I like','I’ve been exploring a few ideas','No, I’m starting from scratch'],
  contact:['','Email','Phone call','WhatsApp'],
  interests:['Culture & history','Food & wine','Nature & wildlife','Beach & relaxation','Adventure & active','Wellness','City breaks','Celebrations & special occasions'],
  status:['New','Reviewing','Contacted','Proposal sent','Booked','Declined','Closed'],
  asset:['Yes','Partial','No']
};

/* ------------------------------------------------------------------ web app */

function doGet() { return json_({ ok:true, service:'saskio-discover', version:'1.0' }); }

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok:false, error:'bad_request' }); }

  // Spam trap: real visitors never see or fill this field. Pretend success, save nothing.
  if (d.website) { log_('spam_trap', 'honeypot filled'); return json_({ ok:true, id:'SD-000000-000' }); }

  const v = validate_(d);
  if (v) { log_('validation', v); return json_({ ok:false, error:'validation', field:v }); }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { log_('busy', String(err)); return json_({ ok:false, error:'busy' }); }

  let id, row, sh;
  try {
    sh = ss_().getSheetByName(CONFIG.SHEET);
    const subCol = AUTO.indexOf('Client submission ID') + 1;
    const last = sh.getLastRow();

    // Re-send of an enquiry that was already saved (e.g. "Try again" after a lost reply): return the same reference.
    if (d.submissionId && last > 1) {
      const from = Math.max(2, last - 199);
      const ids = sh.getRange(from, subCol, last - from + 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (ids[i][0] === d.submissionId) {
          const existing = sh.getRange(from + i, 1).getValue();
          lock.releaseLock();
          return json_({ ok:true, id:existing, duplicate:true });
        }
      }
    }

    const now = new Date();
    const day = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyMMdd');
    let n = 1;
    if (last > 1) {
      const prev = sh.getRange(Math.max(2, last - 299), 1, Math.min(300, last - 1), 1).getValues()
        .map(r => String(r[0])).filter(x => x.indexOf('SD-' + day + '-') === 0);
      n = prev.length + 1;
    }
    id = 'SD-' + day + '-' + ('00' + n).slice(-3);
    const ts = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
    const ages = (d.ages || []).slice(0, d.children).map(a => String(a) === '0' ? '<1' : String(a)).join(', ');

    const values = [
      id, ts, CONFIG.TIMEZONE, s_(d.name), s_(d.email), s_(d.phone), s_(d.contact),
      s_(d.dest), d.dateMode === LISTS.dateMode[0] ? 'Specific' : 'Flexible', ukDate_(d.from), ukDate_(d.to), s_(d.flex), s_(d.duration),
      d.adults, d.children, ages, d.adults + d.children, s_(d.budget),
      s_(d.looking), s_(d.found), (d.interests || []).join('; '), s_(d.notes),
      s_(d.utm_source) || 'unknown', s_(d.utm_medium) || 'unknown', s_(d.utm_campaign), s_(d.utm_content), s_(d.entry), d.fbclid ? 'Yes' : 'No',
      CONFIG.PRIVACY_VERSION, d.marketing ? 'Yes' : 'No', d.marketing ? ts : '',
      'Pending', 'Pending', s_(d.submissionId)
    ];
    // Automated columns are formatted as plain text in setup(); the apostrophe only neutralises formula-like input.
    sh.appendRow(values.map(x => (typeof x === 'string' && /^[=+\-@]/.test(x)) ? "'" + x : x));
    SpreadsheetApp.flush();
    row = sh.getLastRow();
  } catch (err) {
    log_('save_failed', String(err));
    try { lock.releaseLock(); } catch (e2) {}
    return json_({ ok:false, error:'save_failed' });
  }
  lock.releaseLock();

  // Saved. Emails next — a failed email never un-saves the enquiry; it is recorded for follow-up.
  const notifCol = AUTO.indexOf('Internal notification') + 1, ackCol = AUTO.indexOf('Acknowledgement email') + 1;
  const stamp = () => Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM HH:mm');
  try { sendInternal_(id, d); sh.getRange(row, notifCol).setValue('Sent ' + stamp()); }
  catch (err) { sh.getRange(row, notifCol).setValue('FAILED ' + stamp()); log_('notify_failed', id + ' ' + err); }
  try { sendAck_(d); sh.getRange(row, ackCol).setValue('Sent ' + stamp()); }
  catch (err) { sh.getRange(row, ackCol).setValue('FAILED ' + stamp()); log_('ack_failed', id + ' ' + err); }

  return json_({ ok:true, id:id });
}

/* ------------------------------------------------------------------ checks */

function validate_(d) {
  const str = (k, req) => { const x = d[k] == null ? '' : String(d[k]).trim(); if (req && !x) return k; if (LIMITS[k] && x.length > LIMITS[k]) return k; return null; };
  let bad = str('dest', true) || str('name', true) || str('email', true) || str('flex') || str('notes') || str('found') || str('phone');
  if (bad) return bad;
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(d.email))) return 'email';
  if (LISTS.dateMode.indexOf(d.dateMode) < 0) return 'dateMode';
  if (d.dateMode === LISTS.dateMode[0] && !/^\d{4}-\d{2}-\d{2}$/.test(String(d.from || ''))) return 'from';
  if (d.dateMode === LISTS.dateMode[1] && !String(d.flex || '').trim()) return 'flex';
  if (LISTS.duration.indexOf(d.duration || '') < 0) return 'duration';
  if (!(d.adults >= 1 && d.adults <= 12) || !(d.children >= 0 && d.children <= 12)) return 'travellers';
  const ages = (d.ages || []).slice(0, d.children);
  if (ages.length !== d.children || ages.some(a => !(Number(a) >= 0 && Number(a) <= 17) || a === '' || a == null)) return 'ages';
  if (LISTS.budget.indexOf(d.budget) < 0) return 'budget';
  if (LISTS.looking.indexOf(d.looking) < 0) return 'looking';
  if (LISTS.contact.indexOf(d.contact || '') < 0) return 'contact';
  if ((d.contact === 'Phone call' || d.contact === 'WhatsApp') && String(d.phone || '').replace(/\D/g, '').length < 7) return 'phone';
  if ((d.interests || []).some(x => LISTS.interests.indexOf(x) < 0)) return 'interests';
  return null;
}

/* ------------------------------------------------------------------ emails */

// The wordmark image is fetched from the Discover site (and cached) so the script stays small.
// If it can't be fetched, emails fall back to the plain text "Saski’O".
function wordmark_() {
  const cache = CacheService.getScriptCache();
  let b64 = cache.get('wm_b64');
  if (!b64 && CONFIG.SITE_URL.indexOf('http') === 0) {
    try {
      const r = UrlFetchApp.fetch(CONFIG.SITE_URL + 'wordmark.png', { muteHttpExceptions: true });
      if (r.getResponseCode() === 200) { b64 = Utilities.base64Encode(r.getBlob().getBytes()); cache.put('wm_b64', b64, 21600); }
    } catch (e) { log_('wordmark_fetch', String(e)); }
  }
  return b64 ? Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', 'saskio-wordmark.png') : null;
}
let WM_ON = false;
const WM = (h) => WM_ON
  ? '<img src="cid:wm" alt="Saski’O" height="' + h + '" style="height:' + h + 'px;width:auto;vertical-align:baseline;border:0">'
  : '<span style="font-weight:600;letter-spacing:.01em">Saski’O</span>';

function sendAck_(d) {
  const wm = wordmark_(); WM_ON = !!wm;
  const font = "font-family:Manrope,'Helvetica Neue',Arial,sans-serif;";
  const html =
    '<div style="background:#F4F2EF;padding:32px 16px">' +
    '<div style="max-width:520px;margin:0 auto;background:#FBFAF8;border:1px solid #D8D2CC;border-radius:12px;padding:36px 32px;' + font + 'color:#3D3835;font-size:16px;line-height:1.7">' +
    '<p style="margin:0 0 16px">Thank you for contacting ' + WM(13) + '.</p>' +
    '<p style="margin:0 0 16px">I’ll personally review your enquiry and be in touch.</p>' +
    '<p style="margin:0">Speak soon,</p>' +
    '<p style="margin:10px 0 4px;color:#1F1B19">' + WM(24) + '</p>' +
    '<p style="margin:0;color:#1F1B19">x</p>' +
    '<p style="margin:32px 0 0;padding-top:16px;border-top:1px solid #D8D2CC;font-size:12px;line-height:1.6;color:#736C67">' +
    'You’re receiving this because you made an enquiry with ' + WM(10) + '. ' +
    (CONFIG.SITE_URL.indexOf('http') === 0 ? '<a href="' + CONFIG.SITE_URL + '#privacy" style="color:#736C67">Privacy notice</a><br>' : '<br>') +
    '+44 7876 327901 · hello.saskio@gmail.com · Instagram @hello.saskio</p>' +
    '</div></div>';
  const text = 'Thank you for contacting Saski’O.\n\nI’ll personally review your enquiry and be in touch.\n\nSpeak soon,\nSaski’O\nx\n\n' +
    'You’re receiving this because you made an enquiry with Saski’O.\n+44 7876 327901 · hello.saskio@gmail.com · Instagram @hello.saskio';
  MailApp.sendEmail({
    to: String(d.email).trim(), subject: 'Thank you for your enquiry | Saski’O',
    body: text, htmlBody: html, inlineImages: wm ? { wm: wm } : {},
    name: 'Saski’O', replyTo: CONFIG.NOTIFY_TO
  });
}

function sendInternal_(id, d) {
  const rows = [
    ['Reference', id], ['Name', d.name], ['Email', d.email], ['Phone', d.phone || '—'], ['Prefers', d.contact || '—'],
    ['Destination / ideas', d.dest],
    ['When', d.dateMode === LISTS.dateMode[0] ? ukDate_(d.from) + (d.to ? ' – ' + ukDate_(d.to) : '') : d.flex],
    ['Duration', d.duration || '—'], ['Adults (18+)', d.adults],
    ['Children (17 and under)', d.children ? d.children + ' (ages at travel: ' + (d.ages || []).slice(0, d.children).join(', ') + ')' : '0'],
    ['Budget (total)', d.budget], ['Already looking?', d.looking], ['Found a place', d.found || '—'],
    ['Interests', (d.interests || []).join('; ') || '—'], ['Additional details', d.notes || '—'],
    ['Source', (d.utm_source || 'unknown') + ' / ' + (d.utm_medium || 'unknown') + (d.utm_campaign ? ' / ' + d.utm_campaign : '')],
    ['Entry', d.entry || '—'], ['Marketing consent', d.marketing ? 'Yes' : 'No']
  ];
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[c]));
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">' +
    '<p><b>New Discover enquiry received and saved to the register.</b></p>' +
    '<table style="border-collapse:collapse;font-size:14px">' +
    rows.map(r => '<tr><td style="padding:6px 12px 6px 0;color:#666;vertical-align:top;border-bottom:1px solid #eee">' + esc(r[0]) +
      '</td><td style="padding:6px 0;border-bottom:1px solid #eee">' + esc(r[1]) + '</td></tr>').join('') +
    '</table><p><a href="' + ss_().getUrl() + '">Open the enquiry register</a></p>' +
    '<p style="color:#888;font-size:12px">Reply to this email to reply straight to the customer.</p></div>';
  MailApp.sendEmail({
    to: CONFIG.NOTIFY_TO,
    subject: 'New Discover enquiry ' + id + ' — ' + String(d.dest).slice(0, 40) + ' · ' + d.budget,
    body: rows.map(r => r[0] + ': ' + r[1]).join('\n'), htmlBody: html,
    name: 'Saski’O Discover', replyTo: String(d.email).trim()
  });
}

/* ------------------------------------------------------------------ one-off setup */

function setup() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  const get = (name) => ss.getSheetByName(name) || ss.insertSheet(name);

  // Register
  const sh = get(CONFIG.SHEET);
  const headers = AUTO.concat(MANUAL.map(m => m[0]));
  if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setWrap(true).setVerticalAlignment('top');
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  sh.getRange(1, 1, 1, AUTO.length).setBackground('#ECEEF0');
  sh.getRange(1, AUTO.length + 1, 1, 5).setBackground('#FFF1CC');
  sh.getRange(1, AUTO.length + 6, 1, MANUAL.length - 5).setBackground('#FBE3E1');
  sh.getRange(2, 1, sh.getMaxRows() - 1, AUTO.length).setNumberFormat('@');   // automated columns stored as text

  const col = (name) => headers.indexOf(name) + 1;
  const L = (name) => columnLetter_(col(name));
  const rows = sh.getMaxRows() - 1;
  const rng = (name) => sh.getRange(2, col(name), rows, 1);

  // Dropdowns & formats for Victoria's columns
  rng('Status').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(LISTS.status, true).setAllowInvalid(true).build());
  rng('Asset-library match').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(LISTS.asset, true).setAllowInvalid(true).build());
  ['Follow-up date', 'Last contacted', 'Delete after'].forEach(n => rng(n).setNumberFormat('dd/mm/yyyy'));
  ['Est. booking value (£)', 'Net commission (£)', 'Net £ per hour'].forEach(n => rng(n).setNumberFormat('£#,##0.00'));
  ['Gross commission %', 'Saski’O share %'].forEach(n => rng(n).setNumberFormat('0%'));

  // Calculated columns (one formula in the header row covers every new row)
  const A = L('Enquiry ID');
  sh.getRange(1, col('Net commission (£)')).setFormula(
    '={"Net commission (£)";ARRAYFORMULA(IF((' + A + '2:' + A + '="")+(' + L('Est. booking value (£)') + '2:' + L('Est. booking value (£)') + '=""),"",' +
    L('Est. booking value (£)') + '2:' + L('Est. booking value (£)') + '*N(' + L('Gross commission %') + '2:' + L('Gross commission %') + ')*IF(' +
    L('Saski’O share %') + '2:' + L('Saski’O share %') + '="",0.8,' + L('Saski’O share %') + '2:' + L('Saski’O share %') + ')))}');
  sh.getRange(1, col('Net £ per hour')).setFormula(
    '={"Net £ per hour";ARRAYFORMULA(IF((' + L('Net commission (£)') + '2:' + L('Net commission (£)') + '="")+(N(' + L('Projected hours') + '2:' + L('Projected hours') + ')=0),"",' +
    L('Net commission (£)') + '2:' + L('Net commission (£)') + '/' + L('Projected hours') + '2:' + L('Projected hours') + '))}');
  sh.getRange(1, col('vs £25/hr benchmark')).setFormula(
    '={"vs £25/hr benchmark";ARRAYFORMULA(IF(' + L('Net £ per hour') + '2:' + L('Net £ per hour') + '="","",IF(' +
    L('Net £ per hour') + '2:' + L('Net £ per hour') + '>=25,"Above","Below")))}');
  // Delete after: 12 months after last contact, or after submission if never contacted (Submitted is dd/MM/yyyy HH:mm text)
  const B = L('Submitted (UK)'), LC = L('Last contacted');
  sh.getRange(1, col('Delete after')).setFormula(
    '={"Delete after";ARRAYFORMULA(IF(' + A + '2:' + A + '="","",EDATE(IF(' + LC + '2:' + LC + '<>"",' + LC + '2:' + LC +
    ',DATE(MID(' + B + '2:' + B + ',7,4),MID(' + B + '2:' + B + ',4,2),LEFT(' + B + '2:' + B + ',2))),12)))}');

  // Highlights: overdue follow-ups, deletions due, failed emails, below-benchmark
  const cf = [];
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($' + L('Follow-up date') + '2<>"",$' + L('Follow-up date') + '2<TODAY())')
    .setBackground('#FDE2DF').setRanges([rng('Follow-up date')]).build());
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($' + L('Delete after') + '2<>"",$' + L('Delete after') + '2<=TODAY(),$' + L('Status') + '2<>"Booked")')
    .setBackground('#FDE2DF').setRanges([rng('Delete after')]).build());
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('FAILED').setBackground('#FDE2DF').setFontColor('#A1271F')
    .setRanges([rng('Internal notification'), rng('Acknowledgement email')]).build());
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Below').setFontColor('#A1271F').setRanges([rng('vs £25/hr benchmark')]).build());
  sh.setConditionalFormatRules(cf);
  sh.setColumnWidths(1, headers.length, 140);
  [col('Destination / ideas'), col('Found a place — details'), col('Additional details / occasion'), col('Next action')].forEach(c => sh.setColumnWidth(c, 260));

  // Lists
  const li = get(CONFIG.LISTS); li.clear();
  const keys = ['status', 'asset', 'budget', 'looking', 'duration', 'contact', 'interests'];
  keys.forEach((k, i) => { li.getRange(1, i + 1).setValue(k).setFontWeight('bold'); const vals = LISTS[k].filter(x => x !== '').map(x => [x]); li.getRange(2, i + 1, vals.length, 1).setValues(vals); });

  // Data dictionary
  const dd = get(CONFIG.DICT); dd.clear();
  const dict = [['Column', 'Filled by', 'Notes']].concat(
    AUTO.map(h => [h, 'Automation', ''])).concat(MANUAL.map(m => [m[0], m[1] === 'formula' ? 'Formula' : 'Victoria', '']));
  dd.getRange(1, 1, dict.length, 3).setValues(dict); dd.getRange(1, 1, 1, 3).setFontWeight('bold'); dd.setFrozenRows(1);
  dd.getRange(2, 3, 1, 1).setValue('Format SD-YYMMDD-NNN');
  dd.getRange(headers.indexOf('Delete after') + 2, 3).setValue('12 months after last contact (or submission). Delete the row when due unless the enquiry led to a booking.');

  // Log
  const lg = get(CONFIG.LOG); if (lg.getLastRow() === 0) lg.appendRow(['When (UK)', 'Event', 'Detail']);
  lg.getRange(1, 1, 1, 3).setFontWeight('bold');

  const s1 = ss.getSheetByName('Sheet1'); if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  ss.setActiveSheet(sh);
  return 'Setup complete: ' + headers.length + ' columns';
}

/** Run once from the editor to authorise sending email and check it works. Sends one test email to hello.saskio@gmail.com. */
function authoriseAndTest() {
  MailApp.sendEmail({ to: CONFIG.NOTIFY_TO, subject: 'Saski’O Discover — email check', body: 'Discover can send email. Remaining daily quota: ' + MailApp.getRemainingDailyQuota() });
  return 'OK, quota left today: ' + MailApp.getRemainingDailyQuota();
}

/* ------------------------------------------------------------------ helpers */

function ss_() { return SpreadsheetApp.openById(CONFIG.SHEET_ID); }
function s_(x) { return x == null ? '' : String(x).trim(); }
function ukDate_(iso) { if (!iso) return ''; const p = String(iso).split('-'); return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : String(iso); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function columnLetter_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function log_(event, detail) {
  try { ss_().getSheetByName(CONFIG.LOG).appendRow([Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss'), event, String(detail).slice(0, 500)]); } catch (e) {}
}

