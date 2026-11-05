// Digital Pilot OPS backend · VERSION 2026-11-05-X (per-day Role on the COLLAB schedule, break and meal table, Today's shifts tile, voice message in Assign task, Zelyna button moved; everything from W included)
const REPORT_TO = 'birriaz480sj@gmail.com';
const ORDER_TO = 'birriaz480sj@gmail.com';
const SECRET = '';

function getTab_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = tabFind_(ss, name);
  if (!sh) {
    sh = ss.insertSheet(tabName_(name));
    try { tabLook_(sh); } catch (x) {}
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const a = e && e.parameter && e.parameter.action;
  try {
    if (a === 'log') return json_(logGet_(e.parameter.week));
    if (a === 'logweeks') return json_(logWeeks_());
    if (a === 'logpdf') return json_(logPdf_(e.parameter.week));
    if (a === 'tw') return json_(twGet_(e.parameter.week));
    if (a === 'twweeks') return json_(twWeeks_());
    if (a === 'twpdf') return json_(twPdf_(e.parameter.week));
    if (a === 'twhist') return json_(twHistGet_(e.parameter.emp, e.parameter.weeks));
    if (a === 'al') return json_(alGet_());
    if (a === 'rc') return json_(rcGet_(e.parameter.v));
    if (a === 'tv') return json_(tvGet_(e.parameter.v));
    if (a === 'shifts') return json_(shiftsGet_());
    if (a === 'zph') return json_(zphGet_(e.parameter.id));
    if (a === 'mr') return json_(mrList_());
    if (a === 'mrphoto') return json_(mrPhoto_(e.parameter.id));
    if (a === 'vm') return json_(vmList_());
    if (a === 'vmaudio') return json_(vmAudio_(e.parameter.id));
    if (a === 'ping') return json_({ ok: true, t: logStamp_(), stt: !!zelProp_('OPENAI_API_KEY') });
    if (a === 'tr') return json_(tnGet_());
    if (a === 'trplan') return json_(tnPlanGet_());
    if (a === 'inv') return json_(invGet_(e.parameter.cv, e.parameter.loc));
  } catch (err) { return json_({ ok: false, error: String(err) }); }
  return json_(readAll_());
}

function doPost(e) {
  // Zelyna (AI chat) can take several seconds, so it runs BEFORE the shared lock and never blocks the tablets' saves.
  try {
    const b0 = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b0 && b0.action === 'zelyna') return json_(SECRET && b0.key !== SECRET ? { ok: false, error: 'unauthorized' } : zelynaAsk_(b0));
    if (b0 && b0.action === 'trMake') return json_(SECRET && b0.key !== SECRET ? { ok: false, error: 'unauthorized' } : tnMake_(b0));
  } catch (err0) {}
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (SECRET && body.key !== SECRET) return json_({ ok: false, error: 'unauthorized' });
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE MMM d, yyyy h:mm a');

    if (body.action === 'report') {
      const loc = String(body.location || '').slice(0, 50);
      const issue = String(body.issue || '').slice(0, 3000);
      if (!loc || !issue.trim()) return json_({ ok: false, error: 'location and issue are required' });
      const photos = Array.isArray(body.photos) ? body.photos.slice(0, 3) : [];
      const attachments = photos.map((b64, i) =>
        Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'photo-' + (i + 1) + '.jpg'));
      const na = noviMaint_(loc, issue, photos.length, noviRows_('Reports', 4));
      noviMail_({ to: REPORT_TO, report: 'Maintenance Report', period: logFmt_(logNow_().date) + ' at ' + noviTime_(), date: logFmt_(logNow_().date),
        summary: na.summary, obs: na.obs, recs: na.recs, facts: na.facts, attachments: attachments,
        detail: 'Location: ' + loc + '\nTime: ' + stamp + '\nPhotos: ' + photos.length + '\n\nIssue:\n' + issue });
      try {
        getTab_('Reports', ['Date', 'Location', 'Issue', 'Photos']).appendRow([new Date(), loc, issue, photos.length]);
      } catch (x) {}
      return json_({ ok: true });
    }

    if (body.action === 'closeout') {
      const loc = String(body.location || '').slice(0, 50);
      const items = Array.isArray(body.items) ? body.items.map(String) : [];
      const notes = String(body.notes || '').slice(0, 3000);
      if (!loc || !items.length) return json_({ ok: false, error: 'location and items are required' });
      const photos = Array.isArray(body.photos) ? body.photos.slice(0, 3) : [];
      const attachments = photos.map((b64, i) =>
        Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'closeout-' + (i + 1) + '.jpg'));
      const nc = noviClose_(loc, items, notes, photos.length, noviRows_('Close outs', 5));
      noviMail_({ to: REPORT_TO, report: 'Close Out Checklist', period: 'closing on ' + logFmt_(logNow_().date), date: logFmt_(logNow_().date),
        summary: nc.summary, obs: nc.obs, recs: nc.recs, facts: nc.facts, attachments: attachments,
        detail: 'Location: ' + loc + '\nTime: ' + stamp + '\nPhotos: ' + photos.length + '\n\nAll items checked:\n' + items.map(x => '[x] ' + x).join('\n') + '\n\nNotes:\n' + (notes || '(none)') });
      try {
        getTab_('Close outs', ['Date', 'Location', 'Items checked', 'Notes', 'Photos'])
          .appendRow([new Date(), loc, items.length, notes, photos.length]);
      } catch (x) {}
      return json_({ ok: true });
    }

    if (body.action === 'logEntry') return json_(logAppend_(body.entry));
    if (body.action === 'twOp') return json_(twOp_(body));
    if (body.action === 'twDone') return json_(twDone_(body));
    if (body.action === 'twInv') return json_(twInvAdd_(body));
    if (body.action === 'mrAdd') return json_(mrAdd_(body));
    if (body.action === 'collabRep') return json_(collabRep_(body));
    if (body.action === 'mrDone') return json_(mrDone_(body));
    if (body.action === 'mrDel') return json_(mrDel_(body));
    if (body.action === 'empLang') return json_(twEmpLang_(body));
    if (body.action === 'vmAdd') return json_(vmAdd_(body));
    if (body.action === 'vmDel') return json_(vmDel_(body));
    if (body.action === 'vmHeard') return json_(vmHeard_(body));
    if (body.action === 'trSign') return json_(tnSign_(body));
    if (body.action === 'trPlanSet') return json_(tnPlanSet_(body));
    if (body.action === 'twInvSet') return json_(twInvSet_(body));
    if (body.action === 'prepSync') return json_(prepSync_(body));
    if (body.action === 'logEdit') return json_(logEdit_(body.entry));
    if (body.action === 'alSet') return json_(alSet_(body));
    if (body.action === 'fifoLog') return json_(fifoLog_(body.rec));
    if (body.action === 'fifoStatus') return json_(fifoStatus_(body));
    if (body.action === 'logAddItem') return json_(logAddItem_(body.equip, body.food, body.date));

    if (body.action === 'order') {
      const loc = String(body.location || '').slice(0, 50);
      const items = Array.isArray(body.items) ? body.items : [];
      invSendOrder_(loc, items, '');
      return json_({ ok: true });
    }

    if (body.action === 'set') {
      const loc = String(body.location || '').slice(0, 50);
      const name = String(body.item || '').slice(0, 80);
      const code = String(body.code || '').slice(0, 12);
      const count = Math.max(0, Number(body.count) || 0);
      if (!loc || !name) return json_({ ok: false, error: 'location and item required' });
      cntSet_(loc, name, code, count);
      return json_({ ok: true });
    }

    if (body.action === 'clear') {
      const loc = String(body.location || '').slice(0, 50);
      if (loc) cntClear_(loc);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- inventory order: sent from a tablet (Send order) or automatically Sunday + Wednesday at 3 PM ---------- */
function invSendOrder_(loc, items, how) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE MMM d, yyyy h:mm a');
  // The app shows common names; the email and the order logs use the supplier's industry name from the INVENTORY tab.
  try { const ic = logReadJson_(INV_FILE), im = {}; ((ic && ic.items) || []).forEach(function (x) { im[x.c] = x; }); items.forEach(function (i) { const x = im[String(i.code || '').trim().toUpperCase()]; i.app = i.name; if (x && x.ind) i.name = x.ind; }); } catch (x) {}
  const lines = items.map(i => '#' + (i.code || '') + '  ' + i.name + ' x ' + i.count).join('\n');
  const no = noviOrder_(loc, items, ordRows_());
  noviMail_({ to: ORDER_TO, report: 'Inventory Order', period: logFmt_(logNow_().date) + ' at ' + noviTime_(), date: logFmt_(logNow_().date),
    summary: no.summary, obs: no.obs, recs: no.recs, facts: no.facts,
    detail: 'Location: ' + loc + '\nOut of stock items (' + items.length + '):\n\n' + (lines || 'None') + '\n\nSent: ' + stamp + (how ? '\n' + how : '') });
  try { logOrder_(loc, items); } catch (x) {}
  try { addToPurchase_(loc, items); } catch (x) {}
  try { invTabRefresh_(); } catch (x) {}
}

// Sunday + Wednesday at 3 PM (the countdown on the Inventory page): whatever is marked on the tablets is sent automatically,
// exactly like tapping Send order (email, order history, purchase list), and the list is cleared on every tablet.
function invAutoSendTrigger() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (x) { return { ok: false, error: 'busy' }; }
  try {
    const P = PropertiesService.getScriptProperties(), key = logNow_().date;
    if (P.getProperty('INV_AUTO_LAST') === key) return { ok: true, skipped: 'already sent today' };
    P.setProperty('INV_AUTO_LAST', key);
    const d = cntRead_(), cfg = logReadJson_(INV_FILE), pos = {}, out = [], day = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][pDow_(key)];
    ((cfg && cfg.items) || []).forEach(function (i, n) { pos[i.c] = n; });
    Object.keys(d.by).forEach(function (loc) {
      const L = d.by[loc], items = [];
      Object.keys(L).forEach(function (k) { const n = Number(L[k].n) || 0; if (n > 0) items.push({ code: String(L[k].c || ''), name: k, count: n }); });
      if (!items.length) return;
      items.sort(function (a, b) { return (pos[String(a.code).toUpperCase()] === undefined ? 1e9 : pos[String(a.code).toUpperCase()]) - (pos[String(b.code).toUpperCase()] === undefined ? 1e9 : pos[String(b.code).toUpperCase()]); });
      invSendOrder_(loc, items, 'Sent automatically at the ' + day + ' 3 PM order time.');
      cntClear_(loc); out.push(loc + ': ' + items.length + ' items');
    });
    Logger.log('Automatic inventory order: ' + (out.join(', ') || 'nothing was marked, no email sent'));
    return { ok: true, sent: out };
  } finally { try { lock.releaseLock(); } catch (x) {} }
}
function invAutoTrigger_(out) {
  try {
    const mine = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'invAutoSendTrigger'; });
    if (mine.length === 2) return false;
    mine.forEach(function (t) { ScriptApp.deleteTrigger(t); });
    [ScriptApp.WeekDay.SUNDAY, ScriptApp.WeekDay.WEDNESDAY].forEach(function (d) { ScriptApp.newTrigger('invAutoSendTrigger').timeBased().onWeekDay(d).atHour(15).nearMinute(15).inTimezone(LOG_TZ).create(); });
    PropertiesService.getScriptProperties().setProperty('INV_AUTO_TRG', '1');
    if (out) out.push('Automatic inventory order installed (Sunday + Wednesday, 3:00-3:30 PM).');
    return true;
  } catch (x) { if (out) out.push('Automatic inventory order NOT installed: ' + x); return false; }
}

/* ---------- inventory: counts + order history live in Drive files; the INVENTORY tab shows the numbers ---------- */
// There is ONE inventory tab (INVENTORY). The old INVENTORY COUNTS, INVENTORY ORDERS and INVENTORY DASHBOARD tabs
// are moved into two behind-the-scenes Drive files (same folder as the other tablet data) and then removed:
//   inv-counts.json  = how many of each item are on hand  {by:{location:{ITEM NAME:{n, at, c}}}}
//   inv-orders.json  = every item of every order sent    {rows:[[when (ISO), location, item #, item, count], ...]}
const CNT_FILE = 'inv-counts.json', ORD_FILE = 'inv-orders.json', ORD_KEEP_DAYS = 1100;
function invIso_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v) ? v.toISOString() : (v ? String(v) : ''); }
function cntTabRows_() {
  const sh = tabFind_(SpreadsheetApp.getActiveSpreadsheet(), 'Stock');
  return sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
}
// Copies counts from the old tab into the file when the tab's line is newer (or the file does not have it).
function cntMergeTab_(d) {
  let n = 0;
  cntTabRows_().forEach(function (r) {
    if (!r[0] || !r[1]) return;
    const L = d.by[String(r[0])] = d.by[String(r[0])] || {}, k = String(r[1]), at = invIso_(r[3]), cur = L[k];
    if (cur && cur.at && at && cur.at >= at) return;
    if (cur && !at) return;
    L[k] = { n: Math.max(0, Number(r[2]) || 0), at: at, c: String(r[4] === null ? '' : r[4]).trim() }; n++;
  });
  return n;
}
function cntRead_() {
  let d = null; try { d = logReadJson_(CNT_FILE); } catch (x) {}
  if (d && d.by) return d;
  d = { by: {} }; cntMergeTab_(d); logWriteJson_(CNT_FILE, d);
  return d;
}
function cntWrite_(d) { logWriteJson_(CNT_FILE, d); }
function ordTabRows_() {
  const sh = tabFind_(SpreadsheetApp.getActiveSpreadsheet(), 'Orders');
  return sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
}
// Adds the old tab's order lines that the file does not have yet (same time + location + item # + item).
function ordMergeTab_(d) {
  const have = {}; let n = 0;
  d.rows.forEach(function (r) { have[r[0] + '|' + r[1] + '|' + r[2] + '|' + r[3]] = 1; });
  ordTabRows_().forEach(function (r) {
    if (!r[0] || !r[1] || !r[3]) return;
    const row = [invIso_(r[0]), String(r[1]), String(r[2] === null ? '' : r[2]).trim(), String(r[3]), Number(r[4]) || 1], k = row.slice(0, 4).join('|');
    if (have[k]) return; have[k] = 1; d.rows.push(row); n++;
  });
  d.rows.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });
  return n;
}
function ordRead_() {
  let d = null; try { d = logReadJson_(ORD_FILE); } catch (x) {}
  if (d && Array.isArray(d.rows)) return d;
  d = { rows: [] }; ordMergeTab_(d); logWriteJson_(ORD_FILE, d);
  return d;
}
// Order history as [[Date, location, item #, item, count], ...] (oldest first).
function ordRows_() { return ordRead_().rows.map(function (r) { return [new Date(r[0]), r[1], r[2], r[3], r[4]]; }); }

function readAll_() {
  const d = cntRead_(), by = {};
  Object.keys(d.by).forEach(function (loc) {
    Object.keys(d.by[loc]).forEach(function (k) { const n = Number(d.by[loc][k].n) || 0; if (n > 0) { by[loc] = by[loc] || {}; by[loc][k] = n; } });
  });
  return { ok: true, by: by };
}
function cntSet_(loc, name, code, count) {
  const d = cntRead_(), L = d.by[loc] = d.by[loc] || {};
  L[name] = { n: count, at: new Date().toISOString(), c: code };
  cntWrite_(d);
  try { invTabRefresh_(true); } catch (x) {}
}
function cntClear_(loc) {
  const d = cntRead_(), L = d.by[loc] || {}, at = new Date().toISOString();
  Object.keys(L).forEach(function (k) { if (Number(L[k].n) !== 0) { L[k].n = 0; L[k].at = at; } });
  cntWrite_(d);
  try { invTabRefresh_(true); } catch (x) {}
}

function logOrder_(loc, items) {
  const d = ordRead_(), now = new Date().toISOString();
  items.forEach(function (i) { d.rows.push([now, String(loc), String(i.code || '').trim(), String(i.name || ''), Number(i.count) || 1]); });
  const lim = new Date(Date.now() - ORD_KEEP_DAYS * 86400000).toISOString();
  d.rows = d.rows.filter(function (r) { return r[0] >= lim; });
  logWriteJson_(ORD_FILE, d);
}

function dayNum_(d, tz) {
  const p = Utilities.formatDate(d, tz, 'yyyy-MM-dd').split('-');
  return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000;
}

// Pure calculation: rows = [[Date, Location, Code, Item, Count], ...]
function computeStats_(rows, todayNum, dayOf) {
  const g = {};
  rows.forEach(r => {
    if (!r[0] || !r[1] || !r[3]) return;
    const key = r[1] + '|' + (r[2] || r[3]);
    const o = g[key] = g[key] || { loc: r[1], code: r[2], item: r[3], days: {}, qty: 0 };
    o.days[dayOf(r[0])] = 1;
    o.qty += Number(r[4]) || 1;
  });
  const out = Object.keys(g).map(k => {
    const o = g[k];
    const days = Object.keys(o.days).map(Number).sort((a, b) => a - b);
    const n = days.length, last = days[n - 1];
    let avg = null, next = null, left = null, status;
    if (n >= 2) {
      const iv = [];
      for (let i = 1; i < n; i++) iv.push(days[i] - days[i - 1]);
      const use = iv.slice(-5);
      avg = use.reduce((a, b) => a + b, 0) / use.length;
      next = last + Math.round(avg);
      left = next - todayNum;
      const warn = Math.max(2, Math.ceil(avg * 0.25));
      status = left <= 0 ? 'OVERDUE - likely out' : left <= warn ? 'ALMOST - order soon' : 'OK';
    } else {
      status = 'Need more data (1 order)';
    }
    return { loc: o.loc, code: o.code, item: o.item, times: n, qty: o.qty, last: last,
             since: todayNum - last, avg: avg, next: next, left: left, status: status };
  });
  const rank = s => s.indexOf('OVERDUE') === 0 ? 0 : s.indexOf('ALMOST') === 0 ? 1 : s === 'OK' ? 2 : 3;
  out.sort((a, b) => rank(a.status) - rank(b.status) ||
    (a.left === null ? 1e9 : a.left) - (b.left === null ? 1e9 : b.left) ||
    b.times - a.times);
  return out;
}

/* ---------- purchase list: tick the box in column A to close the line ---------- */

const PL = 'Purchase list';
const PL_HEAD = ['Done', 'Date added', 'Location', 'Item #', 'Item', 'Qty', 'Closed at'];

function setupPurchaseSheet_(sh) {
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange(1, 1, 1, PL_HEAD.length).setValues([PL_HEAD])
    .setFontWeight('bold').setBackground('#B30000').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, 998, 1).insertCheckboxes();
  sh.getRange(2, 2, 998, 1).setNumberFormat('MMM d, yyyy');
  sh.getRange(2, 7, 998, 1).setNumberFormat('MMM d, yyyy h:mm a');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2=TRUE')
    .setFontColor('#9E9E9E').setStrikethrough(true).setBackground('#EFEFEF')
    .setRanges([sh.getRange(2, 1, 998, PL_HEAD.length)]).build();
  sh.setConditionalFormatRules([rule]);
  sh.setColumnWidth(1, 60);
  sh.setColumnWidths(2, 1, 110);
  sh.setColumnWidths(3, 1, 110);
  sh.setColumnWidths(4, 1, 70);
  sh.setColumnWidths(5, 1, 300);
  sh.setColumnWidths(6, 1, 60);
  sh.setColumnWidths(7, 1, 160);
}

// Run once from the editor. Uses the Purchase list tab if it exists, else renames the old unused
// Counts tab, else makes a new one.
function setupPurchaseList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = tabFind_(ss, PL);
  if (!sh) {
    const old = tabFind_(ss, 'Counts');
    if (old) { old.setName(tabName_(PL)); sh = old; } else { sh = ss.insertSheet(tabName_(PL)); }
    setupPurchaseSheet_(sh);
  }
}

function getPurchase_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = tabFind_(ss, PL);
  if (!sh) { sh = ss.insertSheet(tabName_(PL)); setupPurchaseSheet_(sh); }
  return sh;
}

// Last row that really has an item (checkbox cells below it do not count)
function lastItemRow_(sh) {
  const v = sh.getRange(1, 5, sh.getMaxRows(), 1).getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]).trim() !== '') return i + 1;
  return 1;
}

// Each order email adds its items here. If the same item is still open for that location, update its qty.
function addToPurchase_(loc, items) {
  const sh = getPurchase_();
  const last = lastItemRow_(sh);
  const data = last > 1 ? sh.getRange(2, 1, last - 1, 7).getValues() : [];
  const now = new Date();
  const fresh = [];
  items.forEach(i => {
    const code = String(i.code || ''), name = String(i.name || ''), qty = Number(i.count) || 1;
    let row = -1;
    for (let r = 0; r < data.length; r++) {
      if (data[r][0] !== true && data[r][2] === loc && (code ? String(data[r][3]) === code : data[r][4] === name)) { row = r; break; }
    }
    if (row >= 0) sh.getRange(row + 2, 6).setValue(qty);
    else fresh.push([false, now, loc, code, name, qty, '']);
  });
  if (fresh.length) {
    const start = lastItemRow_(sh) + 1;
    if (start + fresh.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), fresh.length + 50);
    sh.getRange(start, 1, fresh.length, 7).setValues(fresh);
    sh.getRange(start, 1, fresh.length, 1).insertCheckboxes();
    sh.getRange(start, 2, fresh.length, 1).setNumberFormat('MMM d, yyyy');
    sh.getRange(start, 7, fresh.length, 1).setNumberFormat('MMM d, yyyy h:mm a');
  }
}

// Tick the box in column A: the line greys out, gets a strike-through and a "Closed at" time.
// Untick it to reopen the line.
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== tabName_(PL)) return;
    const r = e.range;
    if (r.getColumn() > 1 || r.getLastColumn() < 1) return;
    for (let row = Math.max(2, r.getRow()); row <= r.getLastRow(); row++) {
      const v = sh.getRange(row, 1).getValue();
      sh.getRange(row, 7).setValue(v === true ? new Date() : '');
    }
  } catch (x) {}
}

function hideClosedPurchases() {
  const sh = getPurchase_();
  const last = lastItemRow_(sh);
  if (last < 2) return;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  vals.forEach((v, i) => { if (v[0] === true) sh.hideRows(i + 2); });
}

function showAllPurchases() {
  const sh = getPurchase_();
  sh.showRows(2, sh.getMaxRows() - 1);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Birriaz')
    .addItem('Refresh inventory numbers (on hand + reorder)', 'rebuildDashboard')
    .addItem('Hide closed purchase items', 'hideClosedPurchases')
    .addItem('Show all purchase items', 'showAllPurchases')
    .addSeparator()
    .addItem('Update tablets with Sheet changes (tasks, prep list, week view, inventory, alarms)', 'twSheetSyncMenu')
    .addItem('Format all tabs (names, colors, red = do not touch)', 'formatSheets')
    .addItem('Build Zelyna knowledge (Sheets + search)', 'zkbBuildMenu')
    .addSeparator()
    .addItem('Make / update COLLAB weekly schedule (keeps a backup)', 'tcBuildMenu')
    .addItem('Space out RECIPES (one blank row between recipes, keeps a backup)', 'rcSpaceOut')
    .addToUi();
}

/* ======================================================================
   FOOD TEMPERATURE LOG - SAN JACINTO
   Data is stored as JSON files in a Google Drive folder (NOT in Google Sheets).
   Records are only ever added. Nothing here deletes or overwrites a reading.
   ====================================================================== */

const LOG_NAME = 'SAN JACINTO';
const LOG_TO = 'birriaz480sj@gmail.com';
const LOG_TZ = 'America/Los_Angeles';
const LOG_FOLDER = 'Digital Pilot OPS - Food Temperature Log';

/*SHARED-START*/
const LOG_CHECKS = ['11:00 AM', '1:00 PM', '3:00 PM', '5:00 PM', '7:00 PM'];
const LOG_MINS = [660, 780, 900, 1020, 1140];
const LOG_EQUIP = ['Steam Table Bay 1', 'Steam Table Bay 2', 'Steam Table Bay 3', 'Food Warmer'];
const LOG_DOW = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
const LOG_DEFAULT_ITEMS = [
  { equip: 'Steam Table Bay 1', food: 'Birria' }, { equip: 'Steam Table Bay 1', food: 'Consomé' },
  { equip: 'Steam Table Bay 2', food: 'Slot 1', flex: true }, { equip: 'Steam Table Bay 2', food: 'Slot 2', flex: true },
  { equip: 'Steam Table Bay 3', food: 'Beans' }, { equip: 'Steam Table Bay 3', food: 'Mexican Rice' },
  { equip: 'Food Warmer', food: 'Birria' }, { equip: 'Food Warmer', food: 'Consomé' }
];
const LOG_FAIL_BELOW = 135;

function logAddDays_(ds, n) {
  const p = ds.split('-');
  return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n)).toISOString().slice(0, 10);
}
function logMonday_(ds) {
  const p = ds.split('-');
  const d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
  return logAddDays_(ds, -((d.getUTCDay() + 6) % 7));
}

// layout = {start:'yyyy-mm-dd', items:[{equip,food,since?}]}; now = {date:'yyyy-mm-dd', min:minutes since midnight}
function logStats_(layout, entries, week, now) {
  const by = {};
  entries.forEach(function (e) {
    const k = e.date + '|' + e.check + '|' + e.equip + '|' + e.food;
    (by[k] = by[k] || []).push(e);
  });
  const items = layout.items.map(function (it, i) { return { it: it, i: i }; })
    .sort(function (a, b) { return (LOG_EQUIP.indexOf(a.it.equip) - LOG_EQUIP.indexOf(b.it.equip)) || (a.i - b.i); })
    .map(function (x) { return x.it; });
  const S = { total: 0, done: 0, passing: 0, missing: 0, pending: 0, failed: 0, actions: 0, empty: 0, discarded: 0, days: [], compliance: null };
  for (let d = 0; d < 7; d++) {
    const date = logAddDays_(week, d), slots = [];
    LOG_CHECKS.forEach(function (ck, ci) {
      items.forEach(function (it) {
        if (date < layout.start || (it.since && date < it.since) || (it.on && it.on !== date)) return;
        const list = by[date + '|' + ck + '|' + it.equip + '|' + it.food] || [];
        const mins = LOG_MINS[ci];
        const abs = date === now.date ? now.min : (date < now.date ? 99999 : -99999);
        let status, req = true;
        if (list.length) { const L = list[list.length - 1]; status = L.empty ? 'empty' : L.result === 'FAIL' ? (L.disd ? 'disc' : 'fail') : 'pass'; }
        else {
          status = abs >= mins + 120 ? 'missing' : abs >= mins - 30 ? 'pending' : 'future';
          if (it.flex) {
            // Flexible spot (Bay 2): empty by default. It only needs a temperature after food was logged in it
            // earlier the same day, and stops needing one once it is logged EMPTY (removed).
            let prior = null;
            for (let k = ci - 1; k >= 0 && !prior; k--) {
              const pl = by[date + '|' + LOG_CHECKS[k] + '|' + it.equip + '|' + it.food];
              if (pl && pl.length) prior = pl[pl.length - 1];
            }
            if (!(prior && !prior.empty)) { req = false; if (status !== 'future') status = 'idle'; }
          }
        }
        if (req || list.length) S.total++;
        if (list.length) S.done++;
        if (status === 'pass' || status === 'empty' || status === 'disc') S.passing++;
        if (status === 'empty') S.empty++;
        if (status === 'disc') S.discarded++;
        if (status === 'missing') S.missing++;
        if (status === 'pending') S.pending++;
        list.forEach(function (e) { if (e.result === 'FAIL') S.failed++; if (e.action) S.actions++; });
        slots.push({ date: date, check: ck, ci: ci, equip: it.equip, food: it.food, flex: !!it.flex, status: status, list: list });
      });
    });
    S.days.push({ date: date, dow: LOG_DOW[d], slots: slots });
  }
  const denom = S.done + S.missing;
  S.compliance = denom ? Math.round(S.passing * 1000 / denom) / 10 : null;
  return S;
}
/*SHARED-END*/
/*TWSHARED-START*/
const TW_SEED = {"emps":["MILTZY","JENNIFER","HUGO","ERIKA","BRISA","ALICIA"],"work":[0,2,3,6],"tasks":[{"id":"t001","row":1,"a":"1 PER WEEK","o":"ELECTRIC FRIER TABLE","en":"ELECTRIC FRYER TABLE","es":"MESA DE LA FREIDORA ELÉCTRICA","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[0],"f":"W1"},{"id":"t002","row":2,"a":"1 PER WEEK","o":"SHELF ABOVE STEAM TABLE","en":"SHELF ABOVE STEAM TABLE","es":"ESTANTE ENCIMA DE LA MESA DE VAPOR","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[2],"f":"W1"},{"id":"t003","row":3,"a":"1 PER WEEK","o":"STEAM TABLE PANS 3 OF THEM","en":"STEAM TABLE PANS 3 OF THEM","es":"CHAROLAS DE LA MESA DE VAPOR (LAS 3)","emp":"MILTZY","src":"sheet","l":[],"fl":"Action not stated. Clean the 3 pans?","d":[3],"f":"W1"},{"id":"t004","row":4,"a":"1 PER WEEK","o":"DISH TOWER AND GRAY BINS","en":"DISH TOWER AND GRAY BINS","es":"TORRE DE PLATOS Y CONTENEDORES GRISES","emp":"JENNIFER","src":"sheet","l":[],"fl":"","d":[0],"f":"W1"},{"id":"t005","row":5,"a":"1 PER WEEK","o":"UNDER ALL FRIDGE AND FREEZER","en":"UNDER ALL FRIDGE AND FREEZER","es":"DEBAJO DE TODOS LOS REFRIGERADORES Y CONGELADORES","emp":"JENNIFER","src":"sheet","l":[],"fl":"","d":[2],"f":"W1"},{"id":"t006","row":6,"a":"1 PER WEEK","o":"2 METAL CART DOLLEYS","en":"2 METAL CART DOLLIES","es":"2 CARRITOS DE METAL","emp":"JENNIFER","src":"sheet","l":[],"fl":"","d":[3],"f":"W1"},{"id":"t007","row":7,"a":"1 PER WEEK","o":"FLOOR DRAINS","en":"FLOOR DRAINS","es":"DRENAJES DEL PISO","emp":"JENNIFER","src":"sheet","l":[],"fl":"","d":[6],"f":"W1"},{"id":"t008","row":8,"a":"1 PER WEEK","o":"KITCHEN FRIDGE AND FREEZER DOOR","en":"KITCHEN FRIDGE AND FREEZER DOOR","es":"PUERTA DEL REFRIGERADOR Y CONGELADOR DE LA COCINA","emp":"JENNIFER","src":"sheet","l":[],"fl":"","d":[0],"f":"W1"},{"id":"t009","row":9,"a":"1 PER WEEK","o":"MESA DE PLANCHAS,","en":"GRIDDLE TABLE","es":"MESA DE PLANCHAS","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[0],"f":"W1"},{"id":"t010","row":10,"a":"1 PER WEEK","o":"CHANGE ALL CONTAINERS AND DEEP CLEAN FRIDGE #6","en":"CHANGE ALL CONTAINERS AND DEEP CLEAN FRIDGE #6","es":"CAMBIAR TODOS LOS RECIPIENTES Y LIMPIAR A FONDO EL REFRIGERADOR #6","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[2],"f":"W1"},{"id":"t011","row":11,"a":"1 PER WEEK","o":"CLEAN FLOOR AND WALL IN FRIDGE #6 LOCATION","en":"CLEAN FLOOR AND WALL IN FRIDGE #6 LOCATION","es":"LIMPIAR PISO Y PARED DONDE ESTÁ EL REFRIGERADOR #6","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[3],"f":"W1"},{"id":"t012","row":12,"a":"1 PER WEEK","o":"SPICE SHELVES","en":"SPICE SHELVES","es":"ESTANTES DE ESPECIAS","emp":"ERIKA","src":"sheet","l":[],"fl":"","d":[0],"f":"W1"},{"id":"t013","row":13,"a":"1 PER WEEK","o":"COOK RED SALSA","en":"COOK RED SALSA","es":"COCINAR SALSA ROJA","emp":"ERIKA","src":"sheet","l":[],"fl":"","d":[2],"f":"W1"},{"id":"t014","row":14,"a":"1 PER WEEK","o":"COOK GREEN SALSA","en":"COOK GREEN SALSA","es":"COCINAR SALSA VERDE","emp":"ERIKA","src":"sheet","l":[],"fl":"","d":[3],"f":"W1"},{"id":"t015","row":15,"a":"1 PER WEEK","o":"PEPSI CO2 AND SODAS","en":"PEPSI CO2 AND SODAS","es":"CO2 DE PEPSI Y REFRESCOS","emp":"ERIKA","src":"sheet","l":[],"fl":"Action not stated. Check, change, or stock?","d":[6],"f":"W1"},{"id":"t016","row":16,"a":"1 PER WEEK","o":"OFFICE, DEEP CLEAN","en":"OFFICE, DEEP CLEAN","es":"OFICINA, LIMPIEZA PROFUNDA","emp":"ERIKA","src":"sheet","l":[],"fl":"","d":[0],"f":"W1"},{"id":"t017","row":17,"a":"1 PER WEEK","o":"OFFICE, CLEAN","en":"OFFICE, CLEAN","es":"OFICINA, LIMPIAR","emp":"ERIKA","src":"sheet","l":[],"fl":"","d":[2],"f":"W1"},{"id":"t018","row":18,"a":"1 PER WEEK","o":"FRIDGE #7 DEEP CLEAN","en":"FRIDGE #7 DEEP CLEAN","es":"LIMPIEZA PROFUNDA DEL REFRIGERADOR #7","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[0],"f":"W1"},{"id":"t019","row":19,"a":"1 PER WEEK","o":"OFFICE, DEBAJO DE HIELERAS BARRER Y MAPEAR","en":"OFFICE, UNDER THE COOLERS: SWEEP AND MOP","es":"OFICINA, DEBAJO DE LAS HIELERAS: BARRER Y MAPEAR","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[2],"f":"W1"},{"id":"t020","row":20,"a":"1 PER WEEK","o":"CLEAN MOP SINK","en":"CLEAN MOP SINK","es":"LIMPIAR LA TARJA DE TRAPEAR","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[3],"f":"W1"},{"id":"t021","row":21,"a":"1 PER WEEK","o":"CLEAN MOP RACK ,WALL AND FLOOR","en":"CLEAN MOP RACK, WALL AND FLOOR","es":"LIMPIAR EL SOPORTE DE TRAPEADORES, PARED Y PISO","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[6],"f":"W1"},{"id":"t022","row":22,"a":"1 PER WEEK","o":"SEAT COVER","en":"SEAT COVER","es":"CUBIERTA DEL ASIENTO","emp":"ALICIA","src":"sheet","l":[],"fl":"Action not stated. Clean it, replace it, or something else?","d":[0],"f":"W1"},{"id":"t023","row":23,"a":"1 PER WEEK","o":"TOILET PAPER IN BATHROOM","en":"TOILET PAPER IN BATHROOM","es":"PAPEL DE BAÑO EN EL BAÑO","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[2],"f":"W1"},{"id":"t024","row":24,"a":"1 PER WEEK","o":"CLEAN BLACK FLOWER TILE","en":"CLEAN BLACK FLOWER TILE","es":"LIMPIAR AZULEJO NEGRO DE FLORES","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[3],"f":"W1"},{"id":"t025","row":25,"a":"1 PER WEEK","o":"CLEAN RAMEN SHELF","en":"CLEAN RAMEN SHELF","es":"LIMPIAR ESTANTE DE RAMEN","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[6],"f":"W1"},{"id":"t026","row":26,"a":"1 PER WEEK","o":"CLEAN FRIDGE #8 AND SANITIZE","en":"CLEAN FRIDGE #8 AND SANITIZE","es":"LIMPIAR Y DESINFECTAR EL REFRIGERADOR #8","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[0],"f":"W1"},{"id":"t027","row":27,"a":"1 PER WEEK","o":"CLEAN UNDER FRIDGE #8","en":"CLEAN UNDER FRIDGE #8","es":"LIMPIAR DEBAJO DEL REFRIGERADOR #8","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[2],"f":"W1"},{"id":"t028","row":28,"a":"1 PER WEEK","o":"CLEAN UNDER SODA RACK","en":"CLEAN UNDER SODA RACK","es":"LIMPIAR DEBAJO DEL ESTANTE DE REFRESCOS","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[3],"f":"W1"},{"id":"t029","row":29,"a":"1 PER WEEK","o":"CLEAN HAND SINK DINING AREA","en":"CLEAN HAND SINK DINING AREA","es":"LIMPIAR LAVAMANOS DEL COMEDOR","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[6],"f":"W1"},{"id":"t030","row":30,"a":"2 PER WEEK","o":"FILL SQUEESE BOTTLES","en":"FILL SQUEEZE BOTTLES","es":"LLENAR BOTES APRETABLES","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[0,3],"f":"W2"},{"id":"t031","row":31,"a":"2 PER WEEK","o":"STAINLESS STEEL WALL BEHIND COOKING EQUIPMENT","en":"STAINLESS STEEL WALL BEHIND COOKING EQUIPMENT","es":"PARED DE ACERO INOXIDABLE DETRÁS DEL EQUIPO DE COCINA","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[2,6],"f":"W2"},{"id":"t032","row":32,"a":"2 PER WEEK","o":"TOWELS","en":"TOWELS","es":"TOALLAS","emp":"ERIKA","src":"sheet","l":[],"fl":"Action not stated. Wash, replace, restock?","d":[0,3],"f":"W2"},{"id":"t033","row":33,"a":"2 PER WEEK","o":"CLEAN BATHROOM","en":"CLEAN BATHROOM","es":"LIMPIAR EL BAÑO","emp":"ERIKA","src":"sheet","l":[],"fl":"","d":[2,6],"f":"W2"},{"id":"t034","row":34,"a":"2 PER WEEK","o":"SOAP DISPENSER","en":"SOAP DISPENSER","es":"DISPENSADOR DE JABÓN","emp":"ALICIA","src":"sheet","l":[],"fl":"Action not stated. Clean it or refill it?","d":[0,3],"f":"W2"},{"id":"t035","row":35,"a":"2 PER WEEK","o":"FLOOR DRANS, AND ACCESORIE PARTS","en":"FLOOR DRAINS AND ACCESSORY PARTS","es":"DRENAJES DEL PISO Y PIEZAS ACCESORIAS","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[2,6],"f":"W2"},{"id":"t036","row":36,"a":"2 PER WEEK","o":"CLEAN CASHIER BOXES","en":"CLEAN CASHIER BOXES","es":"LIMPIAR LAS CAJAS DE LA CAJERA","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[0,3],"f":"W2"},{"id":"t037","row":37,"a":"3 PER WEEK","o":"REPLACE AGUA FRESCA CONTAINERS","en":"REPLACE AGUA FRESCA CONTAINERS","es":"CAMBIAR LOS RECIPIENTES DE AGUA FRESCA","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[2,3,6],"f":"W3"},{"id":"t038","row":38,"a":"ADD TO PREP LIST MONDAY","o":"GREEN ONION","en":"GREEN ONION","es":"CEBOLLA DE CAMBRAY","emp":"MILTZY","src":"sheet","l":["GREEN ONION"],"fl":"","d":[0],"f":"PREP"},{"id":"t039","row":39,"a":"AS NEEDED","o":"WASH DISHES AS NEEDED","en":"WASH DISHES AS NEEDED","es":"LAVAR PLATOS SEGÚN SE NECESITE","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[],"f":"NEED"},{"id":"t040","row":40,"a":"AS NEEDED","o":"ORGANIZE DISHES","en":"ORGANIZE DISHES","es":"ORGANIZAR LOS PLATOS","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[],"f":"NEED"},{"id":"t041","row":41,"a":"AS NEEDED","o":"WASH DISHES AS NEEDED","en":"WASH DISHES AS NEEDED","es":"LAVAR PLATOS SEGÚN SE NECESITE","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"NEED"},{"id":"t042","row":42,"a":"AS NEEDED","o":"ORGANIZE DISHES","en":"ORGANIZE DISHES","es":"ORGANIZAR LOS PLATOS","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"NEED"},{"id":"t043","row":43,"a":"AS NEEDED","o":"SALSAS","en":"SALSAS","es":"SALSAS","emp":"BRISA","src":"sheet","l":[],"fl":"Action not stated. Make, store, or refill salsas?","d":[],"f":"NEED"},{"id":"t044","row":44,"a":"AS NEEDED","o":"STORE INVENTORY, OFFICE, HALL WAY AND BACK AREA AFTER DELIVERY","en":"STORE INVENTORY, OFFICE, HALLWAY AND BACK AREA AFTER DELIVERY","es":"GUARDAR EL INVENTARIO, OFICINA, PASILLO Y ÁREA TRASERA DESPUÉS DE LA ENTREGA","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"NEED"},{"id":"t045","row":45,"a":"AS NEEDED","o":"STORE AWAY RAMEN","en":"STORE AWAY RAMEN","es":"GUARDAR EL RAMEN","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"NEED"},{"id":"t046","row":46,"a":"AS NEEDED","o":"ORGANIZE INVENTORY SHELFS","en":"ORGANIZE INVENTORY SHELVES","es":"ORGANIZAR LOS ESTANTES DEL INVENTARIO","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"NEED"},{"id":"t047","row":47,"a":"AS NEEDED","o":"STORE RICE AND BEANS IN STORAGE CONTAINERS, OFFICE","en":"STORE RICE AND BEANS IN STORAGE CONTAINERS, OFFICE","es":"GUARDAR ARROZ Y FRIJOLES EN RECIPIENTES DE ALMACENAMIENTO, OFICINA","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"NEED"},{"id":"t048","row":48,"a":"AS NEEDED","o":"MAKE FRESAS CON CREMA","en":"MAKE FRESAS CON CREMA","es":"HACER FRESAS CON CREMA","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"NEED"},{"id":"t049","row":49,"a":"DAILY","o":"CLEAN EQUIPMENT KNOBS","en":"CLEAN EQUIPMENT KNOBS","es":"LIMPIAR LAS PERILLAS DEL EQUIPO","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t050","row":50,"a":"DAILY","o":"CLEAN FRIDGE DOORS","en":"CLEAN FRIDGE DOORS","es":"LIMPIAR LAS PUERTAS DE LOS REFRIGERADORES","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t051","row":51,"a":"DAILY","o":"SCRAPE FLAT GRIDDLE","en":"SCRAPE FLAT GRIDDLE","es":"RASPAR LA PLANCHA","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t052","row":52,"a":"DAILY","o":"SCRAPE FLAT GRIDDLE","en":"SCRAPE FLAT GRIDDLE","es":"RASPAR LA PLANCHA","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t053","row":53,"a":"EVERY 2 WEEKS","o":"CLEAN INSIDE FRIDGE 1 AND 2, DEGREASE INSIDE PLATFORM","en":"CLEAN INSIDE FRIDGE 1 AND 2, DEGREASE INSIDE PLATFORM","es":"LIMPIAR POR DENTRO LOS REFRIGERADORES 1 Y 2, DESENGRASAR POR DENTRO LA PLATAFORMA","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[0],"f":"BIW"},{"id":"t054","row":54,"a":"EVERY DAY","o":"CLEAN KITCHEN SURFACES","en":"CLEAN KITCHEN SURFACES","es":"LIMPIAR LAS SUPERFICIES DE LA COCINA","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t055","row":55,"a":"EVERY DAY","o":"CLEAN KITCHEN SPLATTERS","en":"CLEAN KITCHEN SPLATTERS","es":"LIMPIAR SALPICADURAS DE LA COCINA","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t056","row":56,"a":"EVERY DAY","o":"FILL SPRAY BOTTLES","en":"FILL SPRAY BOTTLES","es":"LLENAR BOTES ROCIADORES","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t057","row":57,"a":"EVERY DAY","o":"CLEAN PRINTERS","en":"CLEAN PRINTERS","es":"LIMPIAR LAS IMPRESORAS","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t058","row":58,"a":"EVERY OTHER DAY","o":"CLEAN TICKET HOLDER","en":"CLEAN TICKET HOLDER","es":"LIMPIAR EL PORTA TICKETS","emp":"HUGO","src":"sheet","l":[],"fl":"","d":[],"f":"ODD"},{"id":"t059","row":59,"a":"EVERY OTHER DAY","o":"CHEF TABLE","en":"CHEF TABLE","es":"MESA DEL CHEF","emp":"HUGO","src":"sheet","l":[],"fl":"Action not stated. Clean it?","d":[],"f":"ODD"},{"id":"t060","row":60,"a":"EVERYDAY","o":"TAW OUT FOOD FOR NEXT DAY","en":"THAW OUT FOOD FOR NEXT DAY","es":"DESCONGELAR COMIDA PARA EL DÍA SIGUIENTE","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t061","row":61,"a":"EVERYDAY","o":"ORGANICE FRIDGE 1 AND 2","en":"ORGANIZE FRIDGE 1 AND 2","es":"ORGANIZAR LOS REFRIGERADORES 1 Y 2","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t062","row":62,"a":"EVERYDAY","o":"NACHOS TABLE PLATFORM","en":"NACHOS TABLE PLATFORM","es":"PLATAFORMA DE LA MESA DE NACHOS","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t063","row":63,"a":"EVERYDAY","o":"STEAM TABLE BACK WALL","en":"STEAM TABLE BACK WALL","es":"PARED TRASERA DE LA MESA DE VAPOR","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t064","row":64,"a":"EVERYDAY","o":"CLEAN PRINTERS","en":"CLEAN PRINTERS","es":"LIMPIAR LAS IMPRESORAS","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t065","row":65,"a":"EVERYDAY","o":"SCRAPE FLAT GRIDDLE","en":"SCRAPE FLAT GRIDDLE","es":"RASPAR LA PLANCHA","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t066","row":66,"a":"EVERYDAY","o":"DRAIN SINKS","en":"DRAIN SINKS","es":"DRENAR LAS TARJAS","emp":"MILTZY","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t067","row":67,"a":"EVERYDAY","o":"ORGANIZE DISHES","en":"ORGANIZE DISHES","es":"ORGANIZAR LOS PLATOS","emp":"JENNIFER","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t068","row":68,"a":"EVERYDAY","o":"CLEAN PAPER DISPENSERS","en":"CLEAN PAPER DISPENSERS","es":"LIMPIAR LOS DISPENSADORES DE PAPEL","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t069","row":69,"a":"EVERYDAY","o":"CLEAN METAL SWING DOOR","en":"CLEAN METAL SWING DOOR","es":"LIMPIAR LA PUERTA DE VAIVÉN DE METAL","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t070","row":70,"a":"EVERYDAY","o":"REFILL ROLL PAPER TOWELS","en":"REFILL ROLL PAPER TOWELS","es":"RELLENAR LOS ROLLOS DE TOALLAS DE PAPEL","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t071","row":71,"a":"EVERYDAY","o":"CLEAN PAPER TOWEL UNIT","en":"CLEAN PAPER TOWEL UNIT","es":"LIMPIAR EL DISPENSADOR DE TOALLAS DE PAPEL","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t072","row":72,"a":"EVERYDAY","o":"CHECK DINING AREA SWEEP","en":"CHECK DINING AREA SWEEP","es":"REVISAR EL COMEDOR, BARRER","emp":"BRISA","src":"sheet","l":[],"fl":"Unclear. \"Check dining area, sweep\" or one task?","d":[],"f":"DAY"},{"id":"t073","row":73,"a":"EVERYDAY","o":"CLEAN TABLES","en":"CLEAN TABLES","es":"LIMPIAR LAS MESAS","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t074","row":74,"a":"EVERYDAY","o":"CLEAN HIGH CHAIR","en":"CLEAN HIGH CHAIR","es":"LIMPIAR LA SILLA ALTA","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t075","row":75,"a":"EVERYDAY","o":"CLEAN MICROWAVE","en":"CLEAN MICROWAVE","es":"LIMPIAR EL MICROONDAS","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t076","row":76,"a":"EVERYDAY","o":"CLEAN SODA MACHINE","en":"CLEAN SODA MACHINE","es":"LIMPIAR LA MÁQUINA DE REFRESCOS","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t077","row":77,"a":"EVERYDAY","o":"CHECK HAND SOAP DISP.","en":"CHECK HAND SOAP DISP.","es":"REVISAR EL DISPENSADOR DE JABÓN DE MANOS","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t078","row":78,"a":"EVERYDAY","o":"STOCK SODAS","en":"STOCK SODAS","es":"SURTIR REFRESCOS","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t079","row":79,"a":"EVERYDAY","o":"LIMPIAR SILLA DE NINOS PM","en":"CLEAN KIDS' CHAIR PM","es":"LIMPIAR SILLA DE NIÑOS PM","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t080","row":80,"a":"EVERYDAY","o":"CLEAN VEGGIE SINK","en":"CLEAN VEGGIE SINK","es":"LIMPIAR LA TARJA DE VERDURAS","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t081","row":81,"a":"EVERYDAY","o":"LIMPIAR TELEFONOS","en":"CLEAN PHONES","es":"LIMPIAR TELÉFONOS","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t082","row":82,"a":"EVERYDAY","o":"CLEAN SALSAS STATION WALL","en":"CLEAN SALSA STATION WALL","es":"LIMPIAR LA PARED DE LA ESTACIÓN DE SALSAS","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t083","row":83,"a":"EVERYDAY","o":"DISENFECT PHONES","en":"DISINFECT PHONES","es":"DESINFECTAR LOS TELÉFONOS","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t084","row":84,"a":"EVERYDAY","o":"CLEAN WINDOWS AND MOLDING","en":"CLEAN WINDOWS AND MOLDING","es":"LIMPIAR VENTANAS Y MOLDURAS","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t085","row":85,"a":"EVERYDAY","o":"CHECK TABLE NAPKINS","en":"CHECK TABLE NAPKINS","es":"REVISAR LAS SERVILLETAS DE LAS MESAS","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t086","row":86,"a":"EVERYDAY","o":"CLEAN TABLE NAPKIN DISP.","en":"CLEAN TABLE NAPKIN DISP.","es":"LIMPIAR LOS DISPENSADORES DE SERVILLETAS","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t087","row":87,"a":"EVERYDAY","o":"CLEAN REGISTER/ COUNTER","en":"CLEAN REGISTER/ COUNTER","es":"LIMPIAR LA CAJA REGISTRADORA / MOSTRADOR","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t088","row":88,"a":"EVERYDAY","o":"CHECK BATHROOM /CLEAN/STOCK","en":"CHECK BATHROOM /CLEAN/STOCK","es":"REVISAR EL BAÑO / LIMPIAR / SURTIR","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t089","row":89,"a":"EVERYDAY","o":"CHECK HAND SOAPS","en":"CHECK HAND SOAPS","es":"REVISAR EL JABÓN DE MANOS","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t090","row":90,"a":"EVERYDAY","o":"SANITIZE DOOR/DOOR HANDLES","en":"SANITIZE DOOR/DOOR HANDLES","es":"DESINFECTAR PUERTAS Y MANIJAS","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t091","row":91,"a":"EVERYDAY","o":"CHECK DINING AREA","en":"CHECK DINING AREA","es":"REVISAR EL COMEDOR","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t092","row":92,"a":"EVERYDAY","o":"CLEAN TABLES/HIGH CHAIR","en":"CLEAN TABLES/HIGH CHAIR","es":"LIMPIAR MESAS / SILLA ALTA","emp":"ALICIA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t093","row":93,"a":"EVERYDAY","o":"CHECK BATHROOM /CLEAN/STOCK","en":"CHECK BATHROOM /CLEAN/STOCK","es":"REVISAR EL BAÑO / LIMPIAR / SURTIR","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t094","row":94,"a":"EVERYDAY","o":"CHECK HAND SOAPS","en":"CHECK HAND SOAPS","es":"REVISAR EL JABÓN DE MANOS","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t095","row":95,"a":"EVERYDAY","o":"SANITIZE DOOR/DOOR HANDLES","en":"SANITIZE DOOR/DOOR HANDLES","es":"DESINFECTAR PUERTAS Y MANIJAS","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t096","row":96,"a":"EVERYDAY","o":"CHECK DINING AREA","en":"CHECK DINING AREA","es":"REVISAR EL COMEDOR","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t097","row":97,"a":"EVERYDAY","o":"CLEAN TABLES/HIGH CHAIR","en":"CLEAN TABLES/HIGH CHAIR","es":"LIMPIAR MESAS / SILLA ALTA","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[],"f":"DAY"},{"id":"t098","row":98,"a":"HUGO","o":"TRILL PLATFORM TABLES","en":"TRILL PLATFORM TABLES","es":"TRILL PLATFORM TABLES","emp":"HUGO","src":"sheet","l":[],"fl":"Frequency column says \"HUGO\" and the word \"TRILL\" is unclear (GRILL?). Frequency and wording need your approval.","d":[],"f":"UNK"},{"id":"t099","row":99,"a":"MON, WED, SUN","o":"ORGANIZE DISHES ON THESE DAYS","en":"ORGANIZE DISHES ON THESE DAYS","es":"ORGANIZAR LOS PLATOS EN ESTOS DÍAS","emp":"ERIKA","src":"sheet","l":[],"fl":"","d":[0,2,6],"f":"DAYS"},{"id":"t100","row":100,"a":"MON, WED, SUN","o":"ORGANIZE DISHES ON THESE DAYS","en":"ORGANIZE DISHES ON THESE DAYS","es":"ORGANIZAR LOS PLATOS EN ESTOS DÍAS","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[0,2,6],"f":"DAYS"},{"id":"t101","row":101,"a":"MON, WED, SUN","o":"CLEAN METAL TRASH CAN AFTER TRASH IS REMOVE PM","en":"CLEAN METAL TRASH CAN AFTER TRASH IS REMOVED, PM","es":"LIMPIAR EL BOTE DE BASURA DE METAL DESPUÉS DE SACAR LA BASURA, PM","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[0,2,6],"f":"DAYS"},{"id":"t102","row":102,"a":"PER PREP LIST","o":"DICED \"PICAR\" ASADA","en":"DICED \"PICAR\" ASADA","es":"ASADA PICADA","emp":"MILTZY","src":"sheet","l":["PICAR ASADA"],"fl":"","d":[3],"f":"PREP"},{"id":"t103","row":103,"a":"PER PREP LIST","o":"DICED \"PICAR\" CHICKEN","en":"DICED \"PICAR\" CHICKEN","es":"POLLO PICADO","emp":"MILTZY","src":"sheet","l":["PICAR POLLO PM"],"fl":"","d":[1],"f":"PREP"},{"id":"t104","row":104,"a":"PER PREP LIST","o":"GUACAMOLE","en":"GUACAMOLE","es":"GUACAMOLE","emp":"MILTZY","src":"sheet","l":["GUACAMOLE"],"fl":"","d":[1,6],"f":"PREP"},{"id":"t105","row":105,"a":"PER PREP LIST","o":"PICO","en":"PICO","es":"PICO DE GALLO","emp":"MILTZY","src":"sheet","l":["PICO 1 CONT"],"fl":"","d":[3,6],"f":"PREP"},{"id":"t106","row":106,"a":"PER PREP LIST","o":"TOMATE","en":"TOMATE","es":"TOMATE","emp":"MILTZY","src":"sheet","l":["TOMATE 1 CONT"],"fl":"","d":[3,6],"f":"PREP"},{"id":"t107","row":107,"a":"PER PREP LIST","o":"CILANTRO CHOP \"DICED\"","en":"CILANTRO CHOP \"DICED\"","es":"CILANTRO PICADO","emp":"MILTZY","src":"sheet","l":["PICAR CILANTRO","PICAR CILANTRO TODO"],"fl":"","d":[2,4,5],"f":"PREP"},{"id":"t108","row":108,"a":"PER PREP LIST","o":"GUAC SAUCE","en":"GUAC SAUCE","es":"SALSA DE GUACAMOLE","emp":"MILTZY","src":"sheet","l":["GUAC SAUCE"],"fl":"","d":[6],"f":"PREP"},{"id":"t109","row":109,"a":"PER PREP LIST","o":"PINK SAUCE","en":"PINK SAUCE","es":"SALSA ROSA","emp":"MILTZY","src":"sheet","l":["PINK SAUCE"],"fl":"","d":[6],"f":"PREP"},{"id":"t110","row":110,"a":"PER PREP LIST","o":"FAJITAS","en":"FAJITAS","es":"FAJITAS","emp":"HUGO","src":"sheet","l":["FAJITAS 2-3 CONT"],"fl":"","d":[5],"f":"PREP"},{"id":"t111","row":111,"a":"PER PREP LIST","o":"LETTUCE","en":"LETTUCE","es":"LECHUGA","emp":"HUGO","src":"sheet","l":["LECHUGA 2 CONT"],"fl":"","d":[0,5],"f":"PREP"},{"id":"t112","row":112,"a":"PER PREP LIST","o":"LEMONS","en":"LEMONS","es":"LIMONES","emp":"HUGO","src":"sheet","l":["LIMONES TODO"],"fl":"","d":[0,3],"f":"PREP"},{"id":"t113","row":113,"a":"PER PREP LIST","o":"GRILL  ASADA","en":"GRILL ASADA","es":"ASAR ASADA","emp":"HUGO","src":"sheet","l":["GRILL ASADA"],"fl":"","d":[],"f":"PREP","pn":1},{"id":"t114","row":114,"a":"PER PREP LIST","o":"PURPLE ONION","en":"PURPLE ONION","es":"CEBOLLA MORADA","emp":"ERIKA","src":"sheet","l":["PURPLE ONION TODO"],"fl":"","d":[4],"f":"PREP"},{"id":"t115","row":115,"a":"PER PREP LIST","o":"BIRRIA","en":"BIRRIA","es":"BIRRIA","emp":"ERIKA","src":"sheet","l":["BIRRIA"],"fl":"","d":[1,2,3],"f":"PREP"},{"id":"t116","row":116,"a":"PER PREP LIST","o":"baja sause","en":"BAJA SAUCE","es":"SALSA BAJA","emp":"ERIKA","src":"sheet","l":["BAJA SAUCE"],"fl":"","d":[],"f":"PREP","pn":1},{"id":"t117","row":117,"a":"PER PREP LIST","o":"ROAST CHILES","en":"ROAST CHILES","es":"ASAR CHILES","emp":"ALICIA","src":"sheet","l":["ROAST CHILES"],"fl":"","d":[2],"f":"PREP"},{"id":"t118","row":118,"a":"SUNDAY","o":"FRONT COUNTER TABLE, BOTTOM SHELF AND FLOOR UNDER SHELF","en":"FRONT COUNTER TABLE, BOTTOM SHELF AND FLOOR UNDER SHELF","es":"MESA DEL MOSTRADOR, ESTANTE INFERIOR Y PISO DEBAJO DEL ESTANTE","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[6],"f":"DAYS"},{"id":"t119","row":119,"a":"THURSDAY","o":"TOILET PAPER IN BATHROOM","en":"TOILET PAPER IN BATHROOM","es":"PAPEL DE BAÑO EN EL BAÑO","emp":"BRISA","src":"sheet","l":[],"fl":"","d":[3],"f":"DAYS"},{"id":"t120","row":120,"a":"","o":"DISH TABLE AND GRAY BINS","en":"DISH TABLE AND GRAY BINS","es":"MESA DE PLATOS Y CONTENEDORES GRISES","emp":"JENNIFER","src":"sheet","l":[],"fl":"Frequency column is blank. It may repeat Jennifer's \"DISH TOWER AND GRAY BINS\" (1 PER WEEK). Frequency and wording need your approval.","d":[],"f":"UNK"}]};
function twDayIdx_(ds) { const p = ds.split('-'); return (new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay() + 6) % 7; }
function twDayNum_(ds) { const p = ds.split('-'); return Math.round(Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000); }
function twFind_(cfg, id) { for (let i = 0; i < cfg.tasks.length; i++) if (cfg.tasks[i].id === id) return cfg.tasks[i]; return null; }
// History-safe values: h = [{from:'YYYY-MM-DD', v:value}] in the order they were made. The last change dated on/before `date` wins.
function twAt_(h, date, dflt) { let v = dflt; if (h) for (let i = 0; i < h.length; i++) if (h[i].from <= date) v = h[i].v; return v; }
// Who owns ONE instance: a manager's one-time reassignment, else the collaborator assigned on that date (never rewritten by later changes).
function twEmpAt_(cfg, t, orig) { const o = cfg.iemp && cfg.iemp[t.id + '|' + orig]; if (o) return o; const v = twBase_(cfg, t, orig);
  if (!twAnyPos_(v)) { const e = cfg.employees.filter(function (x) { return x.name === v; })[0]; if (!(e && twHasWk_(e, orig) && !twWorks_(e, orig) && twPosOf_(e.role))) return v; }   // named person working (or no schedule): unchanged
  return twSim_(cfg, t, orig); }
// ---- W: position tasks ("ANY COOK" / "ANY CASHIER" / "ANY DISHES") and same-position backup. Same code on the server and the tablet. ----
// A collaborator's position group comes from the Position typed on the COLLAB tab (CASHIER 2 = Cashier, PM Cook = Cook).
function twPosOf_(role) { const r = String(role || '').toUpperCase(); return /CASH|CAJER/.test(r) ? 'CASHIER' : /COOK|COCIN|GRILL/.test(r) ? 'COOK' : /DISH|LAVA|PLATO|TRASTE/.test(r) ? 'DISHES' : ''; }
function twAnyPos_(v) { const m = /^ANY (CASHIER|COOK|DISHES)$/.exec(String(v || '').trim().toUpperCase()); return m ? m[1] : ''; }
function twWkOf_(e, date) { return e.wkH ? twAt_(e.wkH, date, null) : (e.wk || null); }
function twWorks_(e, date) { const w = twWkOf_(e, date), s = w && w[twDayIdx_(date)]; return !!(s && s.length >= 2); }
function twHasWk_(e, date) { const w = twWkOf_(e, date); return !!(w && w.some(function (s) { return s && s.length >= 2; })); }
// X: the position a person works on that date: the Role picked for that day on the schedule, else the Main position
function twPosDay_(e, date) { const w = twWkOf_(e, date), s = w && w[twDayIdx_(date)]; return (s && s.length >= 3 && twPosOf_(s[2])) || twPosOf_(e.role); }
function twOffAt_(e, date) { return e.offH ? !!twAt_(e.offH, date, e.off ? 1 : 0) : !!e.off; }
function twHash_(s) { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; } return h; }
// Who can take a turn on that date: active people of that position who were added by then. When nobody of the position has hours on the
// schedule yet, the schedule is not used (everyone of the position takes turns); once anyone has hours, only the people working that day.
function twPoolOn_(cfg, pos, date, not) {
  const P = cfg.employees.filter(function (e) { return e.name !== not && twPosDay_(e, date) === pos && !(e.since && date < e.since) && !twOffAt_(e, date); });
  const useWk = P.some(function (e) { return twHasWk_(e, date); });
  return useWk ? P.filter(function (e) { return twWorks_(e, date); }).map(function (e) { return e.name; }) : P.map(function (e) { return e.name; });
}
// Owner before position rules: rotation (Who does it with several names) or the named person.
function twBase_(cfg, t, orig) { const sc = t.sh ? twAt_(t.sh, orig, null) : null, rot = sc ? sc.rot : t.rot, an = sc ? sc.an : t.an; if (rot && rot.length && an) { const L = rot.length; return rot[(((twDayNum_(orig) - twDayNum_(an)) % L) + L) % L]; } return twAt_(t.eh, orig, t.emp); }
// Fair turns: day by day from the start, each turn goes to the person with the fewest turns (then: longest ago; ties: a fixed mixed order per task).
// Nobody of the position working that day: the task waits and goes to the first one working on a later day (up to 2 weeks).
const TW_PPC = typeof WeakMap !== 'undefined' ? new WeakMap() : null;
function twSim_(cfg, t, orig) {
  let C = TW_PPC ? TW_PPC.get(cfg) : null; if (!C) { C = {}; if (TW_PPC) TW_PPC.set(cfg, C); }
  const sig = JSON.stringify(cfg.employees) + '|' + JSON.stringify([t.emp, t.eh, t.rot, t.an, t.sh, t.f, t.d, t.dh, t.ev, t.oh, t.off, t.start, t.date]);
  let M = C[t.id]; if (!M || M.sig !== sig) { M = C[t.id] = { sig: sig, at: '', last: {}, cnt: {}, pick: {} }; }
  if (M.pick[orig] !== undefined) return M.pick[orig];
  const a0 = t.start && t.start > cfg.start ? t.start : cfg.start;
  let d = M.at ? logAddDays_(M.at, 1) : a0;
  if (orig < d) return twPickNow_(cfg, t, orig);   // a date before the start (or not a due day): plain pick
  for (; d <= orig; d = logAddDays_(d, 1)) {
    M.at = d;
    if (!twOn_(cfg, t, d)) continue;
    const base = twBase_(cfg, t, d), pos = twAnyPos_(base); let need = pos, not = '';
    if (!pos) { const e = cfg.employees.filter(function (x) { return x.name === base; })[0]; if (e && twHasWk_(e, d) && !twWorks_(e, d) && twPosOf_(e.role)) { need = twPosOf_(e.role); not = base; } }
    if (!need) { M.pick[d] = base; continue; }
    let who = '';
    for (let k = 0; k < 14 && !who; k++) {
      const day = logAddDays_(d, k), P = twPoolOn_(cfg, need, day, not);
      if (!P.length) continue;
      const known = P.filter(function (p) { return M.cnt[p] !== undefined; }), lo = known.length ? Math.min.apply(null, known.map(function (p) { return M.cnt[p]; })) : 0;
      P.forEach(function (p) { if (M.cnt[p] === undefined) M.cnt[p] = lo; });   // someone new starts level with the others (no catching up)
      P.sort(function (p, q) { const lp = M.last[p] === undefined ? -1 : M.last[p], lq = M.last[q] === undefined ? -1 : M.last[q]; return M.cnt[p] - M.cnt[q] || lp - lq || twHash_(t.id + p) - twHash_(t.id + q); });
      who = P[0]; M.last[who] = twDayNum_(day); M.cnt[who]++;
    }
    M.pick[d] = who || base;
  }
  return M.pick[orig] !== undefined ? M.pick[orig] : twPickNow_(cfg, t, orig);
}
function twPickNow_(cfg, t, orig) {
  const base = twBase_(cfg, t, orig), pos = twAnyPos_(base); let need = pos, not = '';
  if (!pos) { const e = cfg.employees.filter(function (x) { return x.name === base; })[0]; if (!(e && twHasWk_(e, orig) && !twWorks_(e, orig) && twPosOf_(e.role))) return base; need = twPosOf_(e.role); not = base; }
  const P = twPoolOn_(cfg, need, orig, not).sort(function (p, q) { return twHash_(t.id + orig + p) - twHash_(t.id + orig + q); });
  return P[0] || base;
}
// Is a task due on that date (same rules as the week list).
function twOn_(cfg, t, date) {
  const start = cfg.start; if (date < start) return false;
  if (t.oh ? twAt_(t.oh, date, 0) : t.off) return false;
  if (t.start && date < t.start) return false;
  const week = logMonday_(date), d = twDayIdx_(date), w0 = twDayNum_(logMonday_(start));
  const sc = t.sh ? twAt_(t.sh, date, null) : null, T = sc ? Object.assign({}, t, sc) : t;
  const dd = (!sc || sc.init) ? (t.dh ? twAt_(t.dh, date, T.d) : T.d) : T.d;
  let on = false;
  if (T.f === 'ONE') on = T.date === date;
  else if (T.f === 'DAY') on = !(T.rot && T.rot.length && T.an && date < T.an);
  else if (T.f === 'E3' || T.f === 'EV') { const n = T.ev || 3, a = T.an || start; on = date >= a ? (twDayNum_(date) - twDayNum_(a)) % n === 0 : T.f0 === 'DAY' ? true : T.f0 === 'ODD' ? (twDayNum_(date) - twDayNum_(start)) % 2 === 0 : (T.d0 || []).indexOf(d) >= 0; }
  else if (T.f === 'ODD') on = (twDayNum_(date) - twDayNum_(start)) % 2 === 0;
  else if (T.f === 'BIW') on = (dd || []).indexOf(d) >= 0 && Math.round((twDayNum_(week) - w0) / 7) % 2 === 0;
  else if (T.f === 'W1' || T.f === 'W2' || T.f === 'W3' || T.f === 'DAYS' || T.f === 'PREP') on = (dd || []).indexOf(d) >= 0;
  return on;
}

// Snapshot of a task's schedule (used for the schedule history that the Google Sheet writes).
function twSchedOf_(t) { return { f: t.f, d: (t.d || []).slice(), ev: t.ev || 0, an: t.an || '', date: t.date || '', rot: t.rot ? t.rot.slice() : null, d0: t.d0 || null, f0: t.f0 || '' }; }
function twHist0_(t, key, cur) { if (!t[key]) t[key] = [{ from: '0000-00-00', v: cur }]; }
// Completion date = the tablet's date when the box was ticked (offline-safe); the server date is used if the tablet clock looks wrong.
function twDoneDate_(d) { const c = /^\d{4}-\d{2}-\d{2}/.test(d.cts || '') ? d.cts.slice(0, 10) : '', s = (d.sts || '').slice(0, 10); return c && (!s || c <= s) ? c : (s || c); }
// Manager changes. Same code runs on the server and on the tablet (so queued changes show offline).
// o.from (YYYY-MM-DD) = the change only applies from that day on, so earlier days keep their original owner/schedule.
function twApply_(cfg, o) {
  const t = o.tid ? twFind_(cfg, o.tid) : null;
  const okEmp = function (n) { return cfg.employees.some(function (e) { return e.name === n; }); };
  if (o.k === 'assign' && t && (okEmp(o.emp) || twAnyPos_(o.emp))) { if (o.from || t.eh) { twHist0_(t, 'eh', t.emp); t.eh.push({ from: o.from || '0000-00-00', v: o.emp }); } t.emp = o.emp; }
  else if (o.k === 'days' && t) { const nd = (o.d || []).filter(function (x) { return x >= 0 && x <= 6; }).sort(); if (o.from || t.dh) { twHist0_(t, 'dh', (t.d || []).slice()); t.dh.push({ from: o.from || '0000-00-00', v: nd }); } t.d = nd; }
  else if (o.k === 'iassign' && t && /^\d{4}-\d{2}-\d{2}$/.test(o.orig || '') && okEmp(o.emp)) { if (!cfg.iemp) cfg.iemp = {}; const k = o.tid + '|' + o.orig; if (o.emp === twAt_(t.eh, o.orig, t.emp)) delete cfg.iemp[k]; else cfg.iemp[k] = o.emp; }
  else if (o.k === 'excuse' && t && /^\d{4}-\d{2}-\d{2}$/.test(o.orig || '')) { if (!cfg.exc) cfg.exc = {}; const k = o.tid + '|' + o.orig; if (o.on) cfg.exc[k] = { note: String(o.note || '').slice(0, 200), at: String(o.at || '').slice(0, 30) }; else delete cfg.exc[k]; }
  else if (o.k === 'clar' && t) {
    if (o.en) t.en = String(o.en).slice(0, 200); if (o.es) t.es = String(o.es).slice(0, 200);
    if (o.f) { if ((o.f === 'E3' || o.f === 'EV') && t.f !== 'E3' && t.f !== 'EV') { t.d0 = (t.d || []).slice(); t.f0 = t.f; } t.f = o.f; if (o.d) t.d = o.d; if (o.an) t.an = o.an; if (o.ev) t.ev = +o.ev; } t.fl = ''; t.ok = 1;
  }
  else if (o.k === 'move' && t) { const k = o.tid + '|' + o.from; if (o.to && o.to !== o.from) cfg.moves[k] = o.to; else delete cfg.moves[k]; }
  else if (o.k === 'addEmp') { const n = String(o.name || '').trim().toUpperCase().slice(0, 30); if (n && !cfg.employees.some(function (e) { return e.name === n; })) { const ne = { name: n, work: [0, 2, 3, 6] }; if (/^\d{4}-\d{2}-\d{2}$/.test(o.from || '')) ne.since = o.from; cfg.employees.push(ne); } }
  else if (o.k === 'wk') { const e = cfg.employees.filter(function (x) { return x.name === o.emp; })[0]; if (e) { const w = Array.isArray(o.wk) && o.wk.length === 7 ? o.wk.map(function (x) { return Array.isArray(x) && x.length >= 2 ? (x[2] && twPosOf_(x[2]) ? [+x[0], +x[1], twPosOf_(x[2])] : [+x[0], +x[1]]) : null; }) : null; if (o.from || e.wkH) { twHist0_(e, 'wkH', e.wk || null); e.wkH.push({ from: o.from || '0000-00-00', v: w }); } e.wk = w; } }
  else if (o.k === 'work') { const e = cfg.employees.filter(function (x) { return x.name === o.emp; })[0]; if (e) e.work = (o.work || []).sort(); }
  else if (o.k === 'addTask' && o.task && !twFind_(cfg, o.task.id)) {
    const x = o.task; cfg.tasks.push({ id: x.id, src: 'assigned', a: 'ASSIGNED', o: String(x.en).slice(0, 200), en: String(x.en).slice(0, 200), es: String(x.es || x.en).slice(0, 200), emp: x.emp, f: 'ONE', d: [], l: [], fl: '', date: x.date, cr: /^\d{4}-\d{2}-\d{2}$/.test(x.cr || '') ? x.cr : x.date, time: x.time || '', pri: x.pri === 'high' ? 'high' : 'normal' });
  }
  else if (o.k === 'addRot' && o.task && !twFind_(cfg, o.task.id)) {
    // A daily task whose owner rotates: day 0 = rot[0], day 1 = rot[1], ... starting on 'an'. Nothing is scheduled before 'an'.
    const x = o.task; cfg.tasks.push({ id: x.id, src: 'sheet', a: 'DAILY', o: x.en, en: x.en, es: x.es || x.en, emp: x.rot[0], rot: x.rot.slice(), an: x.an, cr: x.an, f: 'DAY', d: [], l: [], fl: '' });
  }
  else if (o.k === 'sched' && t) {
    // From the Google Sheet: new schedule from o.from on. Earlier days keep the old one.
    if (!t.sh) { const s0 = twSchedOf_(t); s0.init = 1; t.sh = [{ from: '0000-00-00', v: s0 }]; }
    const v = { f: o.f, d: (o.d || []).filter(function (x) { return x >= 0 && x <= 6; }).sort(), ev: +o.ev || 0, an: o.an || '', date: o.date || '', rot: (o.rot && o.rot.length) ? o.rot.slice() : null, d0: null, f0: '' };
    t.sh = t.sh.filter(function (h) { return h.from !== (o.from || '0000-00-00') || h.v.init; });
    t.sh.push({ from: o.from || '0000-00-00', v: v });
    t.f = v.f; t.d = v.d; t.ev = v.ev; t.an = v.an; t.date = v.date; if (v.rot) { t.rot = v.rot; t.emp = v.rot[0]; } else delete t.rot;
  }
  else if (o.k === 'addSheet' && o.task && !twFind_(cfg, o.task.id)) {
    const x = o.task, rot = (x.rot && x.rot.length) ? x.rot.slice() : null;
    cfg.tasks.push({ id: x.id, src: 'sheet', a: 'SHEET', o: String(x.en).slice(0, 200), en: String(x.en).slice(0, 200), es: String(x.es || x.en).slice(0, 200), emp: rot ? rot[0] : x.emp, rot: rot || undefined, f: x.f, d: x.d || [], ev: +x.ev || 0, an: x.an || '', date: x.date || '', l: x.l || [], fl: '', cr: x.start, start: x.start });
  }
  else if (o.k === 'empInfo') { const e = cfg.employees.filter(function (x) { return x.name === o.emp; })[0]; if (e) { e.role = String(o.role || '').slice(0, 20); e.shift = String(o.shift || '').slice(0, 20); e.rend = /^[oa]$/i.test(o.rend || '') ? String(o.rend).toLowerCase() : ''; const nf = o.off ? 1 : 0; if (o.from && (e.off ? 1 : 0) !== nf) { twHist0_(e, 'offH', e.off ? 1 : 0); e.offH.push({ from: o.from, v: nf }); } e.off = nf; } }
  else if (o.k === 'link' && t) { t.l = (o.l || []).map(String).slice(0, 5); }
  else if (o.k === 'once' && t && !cfg.once.some(function (x) { return x.tid === o.tid && x.date === o.date; })) cfg.once.push({ tid: o.tid, date: o.date });
  else if (o.k === 'off' && t) { if (o.from || t.oh) { twHist0_(t, 'oh', t.off ? 1 : 0); t.oh.push({ from: o.from || '0000-00-00', v: o.off ? 1 : 0 }); } t.off = o.off ? 1 : 0; }
}
// Every task that is due in one Monday-Sunday week. orig = the day it was first scheduled. due = where it sits now (moved by a manager).
function twInst_(cfg, week) {
  const out = [], start = cfg.start, w0 = twDayNum_(logMonday_(start));
  for (let d = 0; d < 7; d++) {
    const date = logAddDays_(week, d);
    if (date < start) continue;
    cfg.tasks.forEach(function (t) {
      if (twOn_(cfg, t, date)) out.push({ tid: t.id, orig: date, due: date });
    });
  }
  (cfg.once || []).forEach(function (o) { if (logMonday_(o.date) === week && o.date >= start) out.push({ tid: o.tid, orig: o.date, due: o.date, once: true }); });
  out.forEach(function (i) { const m = cfg.moves[i.tid + '|' + i.orig]; if (m && logMonday_(m) === week) i.due = m; });
  return out;
}
function twDoneMap_(comps) {
  const m = {};
  comps.forEach(function (c) { const k = c.tid + '|' + c.orig; if (c.op === 'done') m[k] = { by: c.by, ts: c.sts || c.cts || '', cts: c.cts || '', sts: c.sts || '', id: c.id }; else delete m[k]; });
  return m;
}
// carry = open tasks from earlier weeks [{tid,orig,due}]. final = the week is over (anything still open is missing).
function twStats_(cfg, comps, carry, week, now, final) {
  const dm = twDoneMap_(comps), today = now.date;
  const names = cfg.employees.map(function (e) { return e.name; });
  cfg.tasks.forEach(function (t) { if (names.indexOf(t.emp) < 0 && !twAnyPos_(t.emp)) names.push(t.emp); });
  const items = [];
  function mk(i, isCarry) {
    const t = twFind_(cfg, i.tid); if (!t) return;
    const k = i.tid + '|' + i.orig, d = dm[k], ex = !d && cfg.exc && cfg.exc[k];
    let st = d ? 'done' : ex ? 'excused' : (!final && i.due > today) ? 'future' : (!final && i.due === today) ? 'pending' : 'overdue';
    const cd = d ? twDoneDate_(d) : '', lateDone = !!(cd && cd > i.due);
    const late = st === 'overdue' ? Math.max(0, twDayNum_(today) - twDayNum_(i.due)) : 0;
    items.push({ tid: i.tid, emp: twEmpAt_(cfg, t, i.orig), orig: i.orig, due: i.due, moved: i.due !== i.orig, st: st, carry: isCarry, late: late, by: d ? d.by : '', ts: d ? d.ts : '', cts: d ? d.cts : '', sts: d ? d.sts : '', cd: cd, lateDone: lateDone, daysOver: lateDone ? twDayNum_(cd) - twDayNum_(i.due) : late, cr: t.cr || (t.src === 'assigned' ? t.date : cfg.start), ex: ex ? (ex.note || '') : '', exAt: ex ? (ex.at || '') : '', reas: !!(cfg.iemp && cfg.iemp[k]), src: t.src, pri: t.pri || '', time: t.time || '', link: (t.l || []).length > 0 });
  }
  twInst_(cfg, week).forEach(function (i) { mk(i, false); });
  (carry || []).forEach(function (i) { const k = i.tid + '|' + i.orig; if (!dm[k] && !(cfg.exc && cfg.exc[k])) mk(i, true); });
  const S = { byEmp: {}, order: names, items: items, days: [], store: null };
  function agg(list) {
    const wk = list.filter(function (x) { return !x.carry; }), car = list.filter(function (x) { return x.carry; });
    const done = wk.filter(function (x) { return x.st === 'done'; }).length;
    const miss = wk.filter(function (x) { return x.st === 'overdue'; }).length;
    const pend = wk.filter(function (x) { return x.st === 'pending'; }).length;
    const pendHigh = wk.filter(function (x) { return x.st === 'pending' && x.pri === 'high'; }).length;
    const onTime = wk.filter(function (x) { return x.st === 'done' && !x.lateDone; }).length, lateDone = wk.filter(function (x) { return x.st === 'done' && x.lateDone; }).length;
    const rolled = miss + lateDone + car.length;   // urgent carryover = rolled to a later day at least once
    const due = done + miss + pend;
    const pw = wk.filter(function (x) { return x.link; }), pd = pw.filter(function (x) { return x.st === 'done'; }).length, pdue = pd + pw.filter(function (x) { return x.st === 'overdue'; }).length;
    const aw = wk.filter(function (x) { return x.src === 'assigned'; });
    return { sched: wk.length, done: done, missing: miss, pending: pend, overdue: miss + car.length, urgent: miss + car.length + pendHigh, rolled: rolled, due: due, pct: due ? Math.round(done * 1000 / due) / 10 : null, prepDone: pd, prepDue: pdue, assigned: aw.length, assignedDone: aw.filter(function (x) { return x.st === 'done'; }).length,
      onTime: onTime, lateDone: lateDone, incomplete: miss + pend, carryOpen: car.length, excused: wk.filter(function (x) { return x.st === 'excused'; }).length };
  }
  names.forEach(function (n) {
    const mine = items.filter(function (x) { return x.emp === n; });
    const a = agg(mine);
    a.daily = [];
    for (let d = 0; d < 7; d++) {
      const date = logAddDays_(week, d), di = mine.filter(function (x) { return !x.carry && x.due === date; });
      const dn = di.filter(function (x) { return x.st === 'done'; }).length, mi = di.filter(function (x) { return x.st === 'overdue' || x.st === 'pending'; }).length;
      a.daily.push({ date: date, total: di.length, done: dn, missing: mi, pct: (dn + mi) ? Math.round(dn * 1000 / (dn + mi)) / 10 : null });
    }
    S.byEmp[n] = a;
  });
  S.store = agg(items);
  for (let d = 0; d < 7; d++) { const date = logAddDays_(week, d); S.days.push({ date: date, dow: LOG_DOW[d] }); }
  return S;
}
const TW_MIG = [{"id":"m01","o":{"k":"assign","tid":"t088","emp":"JENNIFER"}},{"id":"m02","o":{"k":"off","tid":"t093","off":1}},{"id":"m03","o":{"k":"assign","tid":"t089","emp":"JENNIFER"}},{"id":"m04","o":{"k":"off","tid":"t094","off":1}},{"id":"m05","o":{"k":"off","tid":"t077","off":1}},{"id":"m06","o":{"k":"assign","tid":"t023","emp":"JENNIFER"}},{"id":"m07","o":{"k":"off","tid":"t119","off":1}},{"id":"m08","o":{"k":"off","tid":"t095","off":1}},{"id":"m09","o":{"k":"off","tid":"t081","off":1}},{"id":"m10","o":{"k":"off","tid":"t091","off":1}},{"id":"m11","o":{"k":"off","tid":"t092","off":1}},{"id":"m12","o":{"k":"off","tid":"t073","off":1}},{"id":"m13","o":{"k":"off","tid":"t074","off":1}},{"id":"m14","o":{"k":"off","tid":"t079","off":1}},{"id":"m15","o":{"k":"off","tid":"t052","off":1}},{"id":"m16","o":{"k":"off","tid":"t040","off":1}},{"id":"m17","o":{"k":"off","tid":"t042","off":1}},{"id":"m18","o":{"k":"off","tid":"t099","off":1}},{"id":"m19","o":{"k":"off","tid":"t100","off":1}},{"id":"m20","o":{"k":"clar","tid":"t003","en":"CLEAN ALL 3 STEAM TABLE PAN INSERTS","es":"LIMPIAR LOS 3 INSERTOS DE LAS CHAROLAS DE LA MESA DE VAPOR"}},{"id":"m21","o":{"k":"clar","tid":"t015","en":"CHECK PEPSI CO2 AND SODAS, CHANGE IF LOW","es":"REVISAR EL CO2 DE PEPSI Y LOS REFRESCOS, CAMBIAR SI ESTÁN BAJOS"}},{"id":"m22","o":{"k":"clar","tid":"t022","en":"CHECK SEAT COVERS, REPLACE IF LOW","es":"REVISAR LOS CUBREASIENTOS, REPONER SI QUEDAN POCOS"}},{"id":"m23","o":{"k":"clar","tid":"t032","en":"WASH TOWELS","es":"LAVAR LAS TOALLAS"}},{"id":"m24","o":{"k":"clar","tid":"t043","en":"REFILL 2 OZ SALSA CUPS AND THAW OUT SALSA CONTAINERS IF NEEDED","es":"RELLENAR VASOS DE SALSA DE 2 OZ Y DESCONGELAR RECIPIENTES DE SALSA SI SE NECESITA"}},{"id":"m25","o":{"k":"clar","tid":"t059","en":"CLEAN CHEF TABLE","es":"LIMPIAR LA MESA DEL CHEF"}},{"id":"m26","o":{"k":"clar","tid":"t072","en":"SWEEP DINING AREA","es":"BARRER EL COMEDOR"}},{"id":"m27","o":{"k":"clar","tid":"t098","en":"CLEAN GRILL PLATFORM TABLES","es":"LIMPIAR LAS MESAS DE LA PLATAFORMA DE LA PARRILLA","f":"ODD","d":[]}},{"id":"m28","o":{"k":"clar","tid":"t120","en":"DISH TABLE AND GRAY BINS","es":"MESA DE PLATOS Y CONTENEDORES GRISES","f":"W1","d":[3]}},{"id":"m29","o":{"k":"clar","tid":"t034","en":"REFILL SOAP DISPENSER","es":"RELLENAR EL DISPENSADOR DE JABÓN"}},{"id":"m30","o":{"k":"clar","tid":"t111","en":"LETTUCE (EVERY 3 DAYS)","es":"LECHUGA (CADA 3 DÍAS)","f":"E3","an":"2026-10-31","d":[]}},{"id":"m31","o":{"k":"off","tid":"t067","off":1,"from":"2026-10-30"}},{"id":"m32","o":{"k":"addRot","task":{"id":"r001","en":"ORGANIZE DISHES (AM SHIFT)","es":"ORGANIZAR LOS PLATOS (TURNO DE LA MAÑANA)","rot":["HUGO","ALICIA","HUGO","JENNIFER"],"an":"2026-10-30"}}},{"id":"m33","o":{"k":"addRot","task":{"id":"r002","en":"ORGANIZE DISHES (PM SHIFT)","es":"ORGANIZAR LOS PLATOS (TURNO DE LA TARDE)","rot":["BRISA","MILTZY","JENNIFER","MILTZY"],"an":"2026-10-30"}}},{"id":"m34","o":{"k":"off","tid":"t072","off":1,"from":"2026-10-31"}},{"id":"m35","o":{"k":"off","tid":"t096","off":1,"from":"2026-10-31"}},{"id":"m36","o":{"k":"off","tid":"t066","off":1,"from":"2026-10-31"}},{"id":"m37","o":{"k":"clar","tid":"t007","en":"WASH FLOOR DRAINS UNIT AND ACCESSORIES","es":"LAVAR LA UNIDAD DE LOS DRENAJES DEL PISO Y SUS ACCESORIOS"}},{"id":"m38","o":{"k":"assign","tid":"t056","emp":"JENNIFER","from":"2026-10-31"}},{"id":"m39","o":{"k":"clar","tid":"t068","en":"CLEAN PAPER DISPENSER UNIT, FIRST DEGREASER THEN SANITIZER","es":"LIMPIAR LA UNIDAD DEL DISPENSADOR DE PAPEL, PRIMERO DESENGRASANTE Y LUEGO DESINFECTANTE"}},{"id":"m40","o":{"k":"clar","tid":"t069","en":"CLEAN METAL SWING DOOR (EVERY 3 DAYS)","es":"LIMPIAR LA PUERTA DE METAL ABATIBLE (CADA 3 DÍAS)","f":"EV","ev":3,"an":"2026-10-31","d":[]}},{"id":"m41","o":{"k":"clar","tid":"t076","en":"CLEAN SODA MACHINE DRAIN, MACHINE OUTER UNIT SURFACE, WALLS BEHIND UNIT AND FLOOR UNDER UNIT","es":"LIMPIAR EL DRENAJE DE LA MÁQUINA DE REFRESCOS, LA SUPERFICIE EXTERIOR DE LA MÁQUINA, LAS PAREDES DETRÁS Y EL PISO DEBAJO"}},{"id":"m42","o":{"k":"clar","tid":"t097","en":"SANITIZE HIGH CHAIR","es":"DESINFECTAR LA SILLA PARA BEBÉ"}},{"id":"m43","o":{"k":"clar","tid":"t082","en":"CLEAN SALSA STATION WALL, CLEAN FLOOR AND ANY OTHER SALSA SPLATTERS","es":"LIMPIAR LA PARED DE LA ESTACIÓN DE SALSAS, LIMPIAR EL PISO Y CUALQUIER OTRA SALPICADURA DE SALSA"}},{"id":"m44","o":{"k":"clar","tid":"t039","en":"WASH DISHES","es":"LAVAR LOS PLATOS","f":"EV","ev":2,"an":"2026-10-31","d":[]}},{"id":"m45","o":{"k":"clar","tid":"t041","en":"WASH DISHES","es":"LAVAR LOS PLATOS","f":"EV","ev":2,"an":"2026-10-31","d":[]}},{"id":"m46","o":{"k":"clar","tid":"t043","f":"EV","ev":2,"an":"2026-10-31","d":[]}},{"id":"m47","o":{"k":"clar","tid":"t044","f":"EV","ev":2,"an":"2026-10-31","d":[]}},{"id":"m48","o":{"k":"clar","tid":"t045","f":"EV","ev":2,"an":"2026-10-31","d":[]}},{"id":"m49","o":{"k":"clar","tid":"t046","f":"EV","ev":2,"an":"2026-10-31","d":[]}},{"id":"m50","o":{"k":"clar","tid":"t047","f":"EV","ev":2,"an":"2026-10-31","d":[]}},{"id":"m51","o":{"k":"clar","tid":"t048","f":"EV","ev":2,"an":"2026-10-31","d":[]}},{"id":"m52","o":{"k":"addEmp","name":"DANIELA"}}];
// Task history across weeks. weeks = [{week, comps}] oldest first. Each scheduled instance appears ONCE, in the week it was due.
function twHist_(cfg, weeks, now) {
  const cur = logMonday_(now.date), items = [], summary = [];
  weeks.forEach(function (w) {
    const st = twStats_(cfg, w.comps || [], [], w.week, now, w.week < cur), by = {};
    st.items.forEach(function (x) { if (x.st !== 'future') { x.week = w.week; items.push(x); } });
    st.order.forEach(function (n) { const a = st.byEmp[n]; by[n] = { pct: a.pct, due: a.due, done: a.done, onTime: a.onTime, lateDone: a.lateDone, missing: a.missing, excused: a.excused }; });
    summary.push({ week: w.week, store: { pct: st.store.pct, due: st.store.due, done: st.store.done, onTime: st.store.onTime, lateDone: st.store.lateDone, missing: st.store.missing }, byEmp: by, order: st.order });
  });
  return { items: items, summary: summary };
}
// Tasks that were missed or finished late more than once in the history window (for schedule / assignment review).
function twRecurring_(cfg, hist, min) {
  const m = {};
  hist.items.forEach(function (x) {
    if (!m[x.tid]) m[x.tid] = { tid: x.tid, n: 0, miss: 0, late: 0, open: 0, emps: {} };
    const r = m[x.tid]; r.n++;
    if (x.st === 'overdue') { r.miss++; r.open++; r.emps[x.emp] = 1; } else if (x.lateDone) { r.late++; r.emps[x.emp] = 1; }
  });
  return Object.keys(m).map(function (k) { const r = m[k], t = twFind_(cfg, k) || {}; r.bad = r.miss + r.late; r.emp = t.emp; r.f = t.f; r.d = t.d || []; r.emps = Object.keys(r.emps); return r; })
    .filter(function (r) { return r.bad >= (min || 2); }).sort(function (a, b) { return b.bad - a.bad || b.open - a.open; });
}
function twMig_(cfg) { if (!cfg.mig) cfg.mig = []; let n = 0; TW_MIG.forEach(function (m) { if (cfg.mig.indexOf(m.id) < 0) { twApply_(cfg, m.o); cfg.mig.push(m.id); n++; } }); return n; }
/*TWSHARED-END*/

/* ---------- storage (Drive JSON files) ---------- */

function logNow_() {
  const n = new Date();
  const hm = Utilities.formatDate(n, LOG_TZ, 'HH:mm').split(':');
  return { date: Utilities.formatDate(n, LOG_TZ, 'yyyy-MM-dd'), min: +hm[0] * 60 + +hm[1] };
}
function logStamp_() { return Utilities.formatDate(new Date(), LOG_TZ, "yyyy-MM-dd'T'HH:mm:ss"); }

function logFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('LOG_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName(LOG_FOLDER);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(LOG_FOLDER);
  props.setProperty('LOG_FOLDER_ID', f.getId());
  return f;
}
function logFile_(name) {
  const it = logFolder_().getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}
function logReadJson_(name) {
  const f = logFile_(name);
  if (!f) return null;
  return JSON.parse(f.getBlob().getDataAsString() || 'null');
}
function logWriteJson_(name, obj) {
  const f = logFile_(name), s = JSON.stringify(obj);
  if (f) f.setContent(s); else logFolder_().createFile(name, s, 'application/json');
}

function logConfig_() {
  let c = logReadJson_('config.json');
  if (!c) {
    c = { start: logNow_().date, items: LOG_DEFAULT_ITEMS.map(function (x) { return x.flex ? { equip: x.equip, food: x.food, flex: true } : { equip: x.equip, food: x.food }; }) };
    logWriteJson_('config.json', c);
  }
  if (!c.items.some(function (x) { return x.flex; })) {   // add the two flexible Bay 2 spots to an older config
    LOG_DEFAULT_ITEMS.filter(function (x) { return x.flex; }).forEach(function (x) { c.items.push({ equip: x.equip, food: x.food, flex: true }); });
    logWriteJson_('config.json', c);
  }
  return c;
}
function logReadWeek_(week, create) {
  const name = 'week-' + week + '.json';
  let d = logReadJson_(name);
  if (!d) {
    d = { week: week, entries: [], report: { status: 'PENDING', attempts: [], sentAt: '' } };
    if (create) logWriteJson_(name, d);
  }
  return d;
}
function logWriteWeek_(week, d) { logWriteJson_('week-' + week + '.json', d); }
function logListWeeks_() {
  const out = [], it = logFolder_().getFiles();
  while (it.hasNext()) {
    const m = /^week-(\d{4}-\d{2}-\d{2})\.json$/.exec(it.next().getName());
    if (m) out.push(m[1]);
  }
  return out.sort();
}

/* ---------- web app actions ---------- */

function logLayoutOut_(c) { return { start: c.start, items: c.items, checks: LOG_CHECKS, equipment: LOG_EQUIP }; }

function logGet_(week) {
  const now = logNow_(), c = logConfig_();
  const wk = logMonday_(/^\d{4}-\d{2}-\d{2}$/.test(week || '') ? week : now.date);
  const d = logReadWeek_(wk, false);
  const last = d.report.attempts.length ? d.report.attempts[d.report.attempts.length - 1] : null;
  return {
    ok: true, now: now, layout: logLayoutOut_(c), week: wk, entries: d.entries,
    report: { status: d.report.status, sentAt: d.report.sentAt, attempts: d.report.attempts.length, lastMsg: last ? last.msg : '' }
  };
}

function logWeeks_() {
  const c = logConfig_(), now = logNow_();
  const weeks = logListWeeks_().reverse().map(function (w) {
    const d = logReadWeek_(w, false), st = logStats_(c, d.entries, w, now);
    return { week: w, entries: d.entries.length, status: d.report.status, compliance: st.compliance, failed: st.failed, missing: st.missing };
  });
  return { ok: true, weeks: weeks };
}

function logAppend_(raw) {
  raw = raw || {};
  const c = logConfig_(), now = logNow_();
  const bad = function (m) { return { ok: false, fatal: true, error: 'invalid: ' + m }; };
  const id = String(raw.id || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) return bad('id');
  const date = String(raw.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return bad('date');
  if (date > logAddDays_(now.date, 1)) return bad('date in the future');
  if (LOG_CHECKS.indexOf(raw.check) < 0) return bad('check time');
  const item = c.items.filter(function (x) { return x.equip === raw.equip && x.food === raw.food && (!x.since || date >= x.since) && (!x.on || x.on === date); })[0];
  if (!item) return bad('equipment or food');
  const empty = raw.empty === true;
  const temp = empty ? null : Number(raw.temp);
  if (!empty && (!isFinite(temp) || temp < 0 || temp > 400 || raw.temp === '' || raw.temp === null)) return bad('temperature');
  const who = String(raw.who || '').trim().slice(0, 30);
  const result = empty ? 'EMPTY' : temp < LOG_FAIL_BELOW ? 'FAIL' : 'PASS';
  const action = String(raw.action || '').trim().slice(0, 500);
  if (result === 'FAIL' && !action) return bad('corrective action required below ' + LOG_FAIL_BELOW + 'F');
  const name = item.flex && !empty ? String(raw.name || '').trim().slice(0, 40) : '';
  if (item.flex && !empty && !name) return bad('food name');
  const wk = logMonday_(date);
  const d = logReadWeek_(wk, true);
  const dup = d.entries.filter(function (e) { return e.id === id; })[0];
  if (dup) return { ok: true, entry: dup, duplicate: true };
  const entry = {
    id: id, restaurant: LOG_NAME, week: wk, date: date, dow: LOG_DOW[(new Date(date + 'T00:00:00Z').getUTCDay() + 6) % 7],
    check: raw.check, equip: raw.equip, food: raw.food, name: name, temp: empty ? null : Math.round(temp * 10) / 10, result: result, empty: empty,
    who: who, action: action, disd: !empty && raw.disd === true,
    kind: raw.kind === 'followup' ? 'followup' : 'initial', ref: String(raw.ref || '').slice(0, 64),
    cts: String(raw.cts || '').slice(0, 30), sts: logStamp_()
  };
  d.entries.push(entry);
  logWriteWeek_(wk, d);
  return { ok: true, entry: entry };
}

function logAddItem_(equip, food, date) {
  const c = logConfig_(), now = logNow_();
  food = String(food || '').trim().slice(0, 40);
  if (LOG_EQUIP.indexOf(equip) < 0 || !food) return { ok: false, fatal: true, error: 'invalid: equipment or food' };
  date = /^\d{4}-\d{2}-\d{2}$/.test(date || '') ? date : now.date;
  // An added food belongs to ONE day only (on: date). It is not asked for on any other day.
  const exists = c.items.some(function (x) { return x.equip === equip && x.food.toLowerCase() === food.toLowerCase() && x.on === date; });
  if (!exists) { c.items.push({ equip: equip, food: food, on: date }); logWriteJson_('config.json', c); }
  return { ok: true, layout: logLayoutOut_(c) };
}

/* ---------- weekly PDF + email ---------- */

const LOG_MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function logFmt_(ds) { const p = ds.split('-'); return LOG_MON[+p[1] - 1] + ' ' + (+p[2]) + ', ' + p[0]; }
function logEsc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function logCompTxt_(st) { return st.compliance === null ? 'n/a' : st.compliance + '%'; }

function logPdfHtml_(layout, st, data, week, now) {
  const BG = { pass: '#CDEBCB', fail: '#F4B6B6', disc: '#F4B6B6', empty: '#DCEBF5', idle: '#FFFFFF', missing: '#FFE14D', pending: '#FFF3B0', future: '#EEEEEE' };
  const sun = logAddDays_(week, 6);
  const th = 'border:1px solid #000;padding:2px 3px;background:#DDD;font-size:8px;text-align:center';
  const td = 'border:1px solid #000;padding:2px 3px;font-size:8px;vertical-align:top';
  let h = '<html><body style="font-family:Arial,Helvetica,sans-serif;color:#000">';
  h += '<div style="font-size:18px;font-weight:bold">FOOD TEMPERATURE LOG &mdash; ' + LOG_NAME + '</div>';
  h += '<div style="font-size:11px;margin:2px 0 6px">Week: Monday ' + logFmt_(week) + ' to Sunday ' + logFmt_(sun) + '</div>';
  h += '<table style="border-collapse:collapse;margin-bottom:8px"><tr>' + [
    ['Scheduled checks', st.total], ['Completed', st.done], ['Missing', st.missing], ['Failed readings', st.failed],
    ['Corrective actions', st.actions], ['Logged empty', st.empty], ['Compliance', logCompTxt_(st)]
  ].map(function (x) { return '<td style="border:1px solid #000;padding:3px 8px;font-size:10px;text-align:center"><b style="font-size:13px">' + x[1] + '</b><br>' + x[0] + '</td>'; }).join('') + '</tr></table>';
  st.days.forEach(function (day) {
    h += '<div style="font-size:11px;font-weight:bold;margin:8px 0 2px">' + day.dow.toUpperCase() + ' ' + logFmt_(day.date) + '</div>';
    if (!day.slots.length) { h += '<div style="font-size:9px;color:#555">Not tracked yet (before the log start date).</div>'; return; }
    h += '<table style="border-collapse:collapse;width:100%"><tr><td style="' + th + '"></td>' +
      LOG_CHECKS.map(function (ck) { return '<td style="' + th + '">DISD?</td><td style="' + th + '"><b>' + ck + '</b></td>'; }).join('') + '</tr>';
    const rows = {};
    day.slots.forEach(function (s) { const k = s.equip + ' | ' + s.food; (rows[k] = rows[k] || []).push(s); });
    let lastEq = '';
    Object.keys(rows).forEach(function (k) {
      const r = rows[k];
      if (r[0].equip !== lastEq) { lastEq = r[0].equip; h += '<tr><td colspan="11" style="' + td + ';background:#F3F3F3;font-weight:bold">' + logEsc_(lastEq.toUpperCase()) + '</td></tr>'; }
      h += '<tr><td style="' + td + ';font-weight:bold">' + logEsc_(r[0].flex ? r[0].food + ' (any food)' : r[0].food) + '</td>';
      LOG_CHECKS.forEach(function (ck) {
        const s = r.filter(function (x) { return x.check === ck; })[0];
        if (!s) { h += '<td style="' + td + '"></td><td style="' + td + '"></td>'; return; }
        const dis = s.list.some(function (e) { return e.disd; });
        let cell = '';
        if (s.list.length) cell = s.list.map(function (e) {
          if (e.empty) return '<div>EMPTY ' + logEsc_(e.who) + '</div>';
          return '<div style="color:' + (e.result === 'FAIL' ? '#B00000' : '#000') + '">' + (e.name ? logEsc_(e.name) + ' ' : '') + (e.result === 'FAIL' ? '! ' : '') + e.temp + '&deg;F ' + logEsc_(e.who) + (e.disd ? ' DISCARDED' : '') + (e.kind === 'followup' ? ' (follow-up)' : '') + '</div>';
        }).join('');
        else if (s.status === 'missing') cell = '<b>MISSING</b>';
        else if (s.status === 'pending') cell = 'pending';
        h += '<td style="' + td + ';text-align:center">' + (dis ? 'X' : '') + '</td><td style="' + td + ';background:' + BG[s.status] + '">' + cell + '</td>';
      });
      h += '</tr>';
    });
    h += '</table>';
  });
  // failed readings + corrective actions
  const fails = [];
  st.days.forEach(function (day) { day.slots.forEach(function (s) { if (s.list.some(function (e) { return e.result === 'FAIL'; })) fails.push(s); }); });
  h += '<div style="font-size:12px;font-weight:bold;margin:12px 0 3px">FAILED READINGS, CORRECTIVE ACTIONS AND FOLLOW-UP TEMPERATURES</div>';
  if (!fails.length) h += '<div style="font-size:9px">No failed readings this week.</div>';
  else {
    h += '<table style="border-collapse:collapse;width:100%"><tr>' + ['Date', 'Check', 'Equipment / Food', 'Reading', 'Who', 'Corrective action', 'Follow-up'].map(function (x) { return '<td style="' + th + '">' + x + '</td>'; }).join('') + '</tr>';
    fails.forEach(function (s) {
      s.list.forEach(function (e, i) {
        if (e.result !== 'FAIL') return;
        const later = s.list.slice(i + 1).map(function (x) { return (x.empty ? 'EMPTY' : x.temp + '&deg;F') + ' ' + logEsc_(x.who); }).join('; ');
        h += '<tr>' + [logFmt_(s.date), s.check, logEsc_(s.equip + ' / ' + (e.name || s.food)), e.temp + '&deg;F' + (e.disd ? ' (DISCARDED)' : ''), logEsc_(e.who), logEsc_(e.action), later || (e.disd ? 'food discarded' : 'none recorded')]
          .map(function (x) { return '<td style="' + td + '">' + x + '</td>'; }).join('') + '</tr>';
      });
    });
    h += '</table>';
  }
  // missing
  h += '<div style="font-size:12px;font-weight:bold;margin:12px 0 3px">MISSING OR INCOMPLETE CHECKS</div>';
  let any = false;
  st.days.forEach(function (day) {
    LOG_CHECKS.forEach(function (ck) {
      const m = day.slots.filter(function (s) { return s.check === ck && s.status === 'missing'; });
      if (m.length) { any = true; h += '<div style="font-size:9px">' + day.dow.substr(0, 3) + ' ' + logFmt_(day.date) + ' ' + ck + ': ' + m.map(function (s) { return logEsc_(s.food + ' (' + s.equip + ')'); }).join(', ') + '</div>'; }
    });
  });
  if (!any) h += '<div style="font-size:9px">None.</div>';
  h += '<div style="font-size:9px;margin-top:10px">Compliance = checks completed &divide; (checks completed + checks missed). A check is completed when it has a reading of ' + LOG_FAIL_BELOW + '&deg;F or higher, is logged EMPTY (no food in that spot), or was below ' + LOG_FAIL_BELOW + '&deg;F and the food was discarded (DISD?). A check counts as missed 2 hours after its scheduled time.</div>';
  h += '<div style="font-size:9px;font-weight:bold;margin-top:10px">TASTE ALL FOOD BEFORE STARTING SHIFT. FOOD MUST BE ABOVE 165&deg;F BEFORE PUTTING IT IN THE STEAM TABLE. REHEAT EVERY 2 HRS IF NO ROTATION IS MADE. IF UNDER 135&deg;F INCREASE HEAT DIAL, ADD CONSOME, AND BRING ABOVE 135&deg;F. DISCARD AFTER 2 HRS IF FOOD WAS UNDER 135&deg;F.</div>';
  h += '<div style="font-size:9px;font-style:italic;margin-top:4px">PRUEBE TODA LA COMIDA ANTES DE COMENZAR SU TURNO. ASEG&Uacute;RESE DE QUE LA COMIDA EST&Eacute; A M&Aacute;S DE 165&deg;F ANTES DE COLOCARLA EN LA MESA DE VAPOR. RECALIENTE CADA 2 HORAS SI NO SE ROTA. SI LA TEMPERATURA ES MENOR A 135&deg;F, AUMENTE EL CONTROL DE CALENTAMIENTO, AGREGUE CONSOM&Eacute; Y LLEVE LA TEMPERATURA A M&Aacute;S DE 135&deg;F. DESECHAR DESPU&Eacute;S DE 2 HORAS SI EL ALIMENTO ESTABA A MENOS DE 135&deg;F.</div>';
  h += '<div style="font-size:8px;color:#555;margin-top:6px">Generated ' + logStamp_().replace('T', ' ') + ' Pacific time. Digital Pilot OPS by Birriaz Restaurants.</div></body></html>';
  return h;
}

function logWeekPdf_(week) {
  const c = logConfig_(), now = logNow_(), d = logReadWeek_(week, false);
  const st = logStats_(c, d.entries, week, now);
  const name = LOG_NAME + ' - Food Temperature Log - ' + week + ' to ' + logAddDays_(week, 6);
  const blob = Utilities.newBlob(logPdfHtml_(c, st, d, week, now), 'text/html', name + '.html').getAs('application/pdf').setName(name + '.pdf');
  return { blob: blob, st: st, name: name };
}

function logPdf_(week) {
  const wk = logMonday_(/^\d{4}-\d{2}-\d{2}$/.test(week || '') ? week : logNow_().date);
  const r = logWeekPdf_(wk);
  return { ok: true, name: r.name + '.pdf', b64: Utilities.base64Encode(r.blob.getBytes()) };
}

// Sends the report email for one week. test=true sends a copy labelled TEST and records nothing.
function logSendWeek_(week, test) {
  const r = logWeekPdf_(week), st = r.st, sun = logAddDays_(week, 6);
  const fm = function (ds) { return logFmt_(ds).toUpperCase(); };
  const pw = logAddDays_(week, -7), c0 = logConfig_(), prev = pw >= logMonday_(c0.start) ? logStats_(c0, logReadWeek_(pw, false).entries, pw, logNow_()) : null;
  const nt = noviTemp_(week, st, logReadWeek_(week, false).entries, prev);
  const detail = LOG_NAME + ' weekly food temperature log\n' +
    'Week: Monday ' + logFmt_(week) + ' to Sunday ' + logFmt_(sun) + '\n\n' +
    'Total completed checks: ' + st.done + ' of ' + st.total + '\n' +
    'Total missed checks: ' + st.missing + '\n' +
    'Checks logged empty (no food): ' + st.empty + '\n' +
    'Total failed readings: ' + st.failed + '\n' +
    'Corrective actions recorded: ' + st.actions + '\n' +
    'Overall weekly compliance: ' + logCompTxt_(st) + '\n\n' +
    'The full printable log is attached as a PDF.';
  const res = { t: Date.now(), at: logStamp_(), ok: false, msg: '' };
  try {
    noviMail_({ to: LOG_TO, test: test, report: 'Weekly Food Temperature Log', period: 'the week of Monday ' + logFmt_(week) + ' to Sunday ' + logFmt_(sun), date: noviRange_(week, sun),
      summary: nt.summary, obs: nt.obs, recs: nt.recs, facts: nt.facts, detail: detail, attachments: [r.blob] });
    res.ok = true;
    res.msg = 'Accepted by Gmail/MailApp with no error. Daily send quota left: ' + MailApp.getRemainingDailyQuota();
  } catch (err) { res.msg = 'FAILED: ' + String(err); }
  if (!test) {
    const d = logReadWeek_(week, true);
    d.report.attempts.push(res);
    if (res.ok) { d.report.status = 'SENT'; d.report.sentAt = res.at; }
    logWriteWeek_(week, d);
  }
  return res;
}

// Runs every 30 minutes and on Sunday 9 PM. Sends any finished week that has not been sent yet. Never sends twice.
function logProcess_() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const c = logConfig_(), now = logNow_();
    let w = logMonday_(c.start), n = 0;
    const cur = logMonday_(now.date);
    while (w <= cur && n++ < 520) { logReadWeek_(w, true); w = logAddDays_(w, 7); }
    logListWeeks_().forEach(function (week) {
      const d = logReadWeek_(week, false);
      if (d.report.status === 'SENT') return;
      const sun = logAddDays_(week, 6);
      const due = now.date > sun || (now.date === sun && now.min >= 1260);
      if (!due) return;
      const last = d.report.attempts[d.report.attempts.length - 1];
      if (last && Date.now() - last.t < 25 * 60000) return;
      logSendWeek_(week, false);
    });
  } finally { try { lock.releaseLock(); } catch (x) {} }
}
function logWeeklyTrigger() { logProcess_(); twProcess_(); }
// Sunday 11:00 PM Pacific: the weekly collaborator task report.
function twWeeklyTrigger() { twProcess_(); }
// Every 30 minutes: retries reports AND makes sure today's 12:05 AM task rollover happened (catch-up if Google delayed it).
function logRetryTrigger() { logProcess_(); twProcess_(); twRollover_(false); prepRoll_(); try { twSheetSync_(); } catch (x) {} }
// Every day at 12:05 AM Pacific (Google runs it within about 15 minutes of that time).
function twDailyRolloverTrigger() { twRollover_(false); prepRoll_(); }

// Run ONCE from the editor (approve the Drive + Gmail permissions). Creates the folder, the config and both triggers.
function setupLogTemps() {
  logConfig_();
  twConfig_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const h = t.getHandlerFunction();
    if (h === 'logWeeklyTrigger' || h === 'logRetryTrigger' || h === 'twDailyRolloverTrigger' || h === 'twWeeklyTrigger') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('logWeeklyTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(21).nearMinute(0).inTimezone(LOG_TZ).create();
  ScriptApp.newTrigger('logRetryTrigger').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('twDailyRolloverTrigger').timeBased().everyDays(1).atHour(0).nearMinute(5).inTimezone(LOG_TZ).create();
  ScriptApp.newTrigger('twWeeklyTrigger').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).nearMinute(0).inTimezone(LOG_TZ).create();
}
// Same as setupLogTemps (safe to run again: it replaces the triggers, it never duplicates them or touches saved data).
function setupTaskAccountability() { setupLogTemps(); twRollover_(true); }

// Run from the editor to test the PDF email. Sends the CURRENT week with TEST in the subject. Records nothing.
function sendTestLogReport() {
  const res = logSendWeek_(logMonday_(logNow_().date), true);
  Logger.log(JSON.stringify(res));
}

// Optional: run from the editor to change the first day the log counts, e.g. setLogStartDate('2026-11-02')
function setLogStartDate(ds) {
  const c = logConfig_();
  c.start = ds;
  logWriteJson_('config.json', c);
}

/* ===================== COLLABORATOR TASKS (Open list) ===================== */
// Stored as JSON files in the same Drive folder as the temperature log. No Google Sheets. Nothing is ever deleted automatically.
const TW_CFG = 'tasks-config.json', TW_INVF = 'inv-extra.json';
const TW_PIN = '4535';   // manager code for the Collaborator task list (checked here, on the server)

function twConfig_() {
  let c = logReadJson_(TW_CFG);
  if (!c) {
    c = { v: 1, start: logNow_().date, employees: TW_SEED.emps.map(function (n) { return { name: n, work: TW_SEED.work.slice() }; }),
      tasks: JSON.parse(JSON.stringify(TW_SEED.tasks)), moves: {}, once: [], applied: [] };
    logWriteJson_(TW_CFG, c);
  }
  if (twMig_(c)) logWriteJson_(TW_CFG, c);   // owner-approved changes, applied once
  return c;
}
function twWeekFile_(w) { return 'tw-' + w + '.json'; }
function twReadWeek_(week, create) {
  let d = logReadJson_(twWeekFile_(week));
  if (!d) { d = { week: week, comps: [], report: { status: 'PENDING', attempts: [] } }; if (create) logWriteJson_(twWeekFile_(week), d); }
  if (!d.report) d.report = { status: 'PENDING', attempts: [] };
  return d;
}
// Open (not completed) tasks from weeks before `week`, so a missed task never disappears when a new week starts.
function twCarry_(cfg, week) {
  const out = []; let w = logMonday_(cfg.start), n = 0;
  while (w < week && n++ < 120) {
    const d = twReadWeek_(w, false), dm = twDoneMap_(d.comps);
    twInst_(cfg, w).forEach(function (i) { const k = i.tid + '|' + i.orig; if (!dm[k] && !(cfg.exc && cfg.exc[k])) out.push({ tid: i.tid, orig: i.orig, due: i.due }); });
    w = logAddDays_(w, 7);
  }
  return out;
}
function twInvRead_() { return logReadJson_(TW_INVF) || { pending: [], extra: [] }; }
// Language of one collaborator's profile on the tablets (EN / ES). Only the profile uses it, not the rest of the app.
function twEmpLang_(raw) {
  const n = String((raw && raw.name) || '').trim().toUpperCase(), lg = String((raw && raw.lang) || '');
  if (['en', 'es', ''].indexOf(lg) < 0) return { ok: false, error: 'bad language' };
  const cfg = twConfig_(), e = cfg.employees.filter(function (x) { return x.name === n; })[0];
  if (!e) return { ok: false, error: 'Collaborator not found.' };
  if (lg) e.lang = lg; else delete e.lang;
  logWriteJson_(TW_CFG, cfg);
  return { ok: true };
}
function twGet_(week) {
  const now = logNow_(), cfg = twConfig_();
  const wk = logMonday_(/^\d{4}-\d{2}-\d{2}$/.test(week || '') ? week : now.date);
  const d = twReadWeek_(wk, false);
  const cfgOut = JSON.parse(JSON.stringify(cfg)); delete cfgOut.applied; delete cfgOut.log;
  const roll = {}; Object.keys(d.roll || {}).forEach(function (k) { const r = d.roll[k]; roll[k] = { at: r.at, n: r.n, keys: (r.items || []).map(function (x) { return x.tid + '|' + x.orig; }) }; });
  return { ok: true, now: now, week: wk, cfg: cfgOut, comps: d.comps, report: d.report, carry: twCarry_(cfg, wk), inv: twInvRead_(), roll: roll };
}
function twWeeks_() {
  const cfg = twConfig_(), now = logNow_(), out = []; let w = logMonday_(cfg.start), n = 0;
  while (w <= logMonday_(now.date) && n++ < 520) {
    const d = twReadWeek_(w, false), st = twStats_(cfg, d.comps, twCarry_(cfg, w), w, now, w < logMonday_(now.date));
    const be = {}; st.order.forEach(function (n) { be[n] = st.byEmp[n].pct; });
    out.push({ week: w, pct: st.store.pct, done: st.store.done, missing: st.store.missing, onTime: st.store.onTime, lateDone: st.store.lateDone, status: d.report.status, byEmp: be, snap: d.report.snap || null }); w = logAddDays_(w, 7);
  }
  return { ok: true, weeks: out.reverse() };
}
function twOp_(raw) {
  const o = raw && raw.o;
  if (!raw || !raw.id || !o || ['assign', 'days', 'clar', 'move', 'addEmp', 'work', 'addTask', 'once', 'off', 'iassign', 'excuse'].indexOf(o.k) < 0) return { ok: false, error: 'bad change' };
  if (o.k !== 'once' && !TW_MGR_EDITS) return { ok: false, error: 'Tasks are changed on the Tasks tab of the Google Sheet now.' };
  if (o.k !== 'once' && String(raw.pin || '') !== TW_PIN) return { ok: false, error: 'wrong manager code' };
  const cfg = twConfig_();
  if (cfg.applied.indexOf(raw.id) >= 0) return { ok: true, dup: true };
  // History protection: an older tablet page that sends no date still only changes today and later.
  if ((o.k === 'assign' || o.k === 'days' || o.k === 'off') && !/^\d{4}-\d{2}-\d{2}$/.test(o.from || '')) o.from = logNow_().date;
  if (o.k === 'excuse' && !o.at) o.at = logStamp_();
  if (o.k === 'addTask' && o.task && !o.task.cr) o.task.cr = logNow_().date;
  twApply_(cfg, o);
  if (!cfg.log) cfg.log = [];
  cfg.log.push({ id: String(raw.id).slice(0, 40), k: o.k, tid: o.tid || '', orig: o.orig || o.from || '', emp: o.emp || '', note: String(o.note || '').slice(0, 120), at: logStamp_() });   // manager audit trail (never trimmed)
  cfg.applied.push(raw.id); if (cfg.applied.length > 500) cfg.applied = cfg.applied.slice(-400);
  logWriteJson_(TW_CFG, cfg);
  return { ok: true };
}
function twDone_(raw) {
  if (!raw || !raw.id || !/^\d{4}-\d{2}-\d{2}$/.test(raw.orig || '') || (raw.op !== 'done' && raw.op !== 'undo')) return { ok: false, error: 'bad entry' };
  const cfg = twConfig_();
  if (!twFind_(cfg, raw.tid)) return { ok: false, error: 'unknown task' };
  const wk = logMonday_(raw.orig), d = twReadWeek_(wk, true);
  if (d.comps.some(function (c) { return c.id === raw.id; })) return { ok: true, dup: true };
  const c = { id: String(raw.id).slice(0, 40), tid: raw.tid, orig: raw.orig, op: raw.op, by: String(raw.by || '').slice(0, 30), cts: String(raw.cts || '').slice(0, 30), sts: logStamp_() };
  d.comps.push(c); logWriteJson_(twWeekFile_(wk), d);
  return { ok: true, comp: c };
}
function twNorm_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9ÁÉÍÓÚÑ]+/g, ' ').trim(); }
function twInvAdd_(raw) {
  const nm = String((raw && raw.name) || '').trim().slice(0, 60);
  if (!raw || !raw.id || !nm) return { ok: false, error: 'bad item' };
  const f = twInvRead_();
  if (f.pending.some(function (p) { return p.id === raw.id; })) return { ok: true, dup: true };
  const ex = f.pending.concat(f.extra).filter(function (p) { return twNorm_(p.name || p.n) === twNorm_(nm) && p.status !== 'rejected'; })[0];
  if (ex) return { ok: true, dup: true, existing: ex };
  f.pending.push({ id: String(raw.id).slice(0, 40), name: nm, cat: String(raw.cat || '').slice(0, 30), qty: String(raw.qty || '').slice(0, 20), note: String(raw.note || '').slice(0, 200), status: 'pending', cts: String(raw.cts || '').slice(0, 30), sts: logStamp_() });
  logWriteJson_(TW_INVF, f);
  // Also listed on the INVENTORY tab so a manager can give it an Item # and set Active to Yes.
  try { const sh = tabFind_(SpreadsheetApp.getActiveSpreadsheet(), INV_TAB); if (sh) sh.appendRow(['', nm.toUpperCase(), nm.toUpperCase(), INV_CATS.indexOf(String(raw.cat || '').toUpperCase()) >= 0 ? String(raw.cat).toUpperCase() : '', '', 'No', 'Requested on a tablet ' + twFmtTs_(logStamp_()) + (raw.qty ? ' (' + String(raw.qty).slice(0, 20) + ')' : '') + (raw.note ? ': ' + String(raw.note).slice(0, 100) : '') + '. Give it an Item #, set Active to Yes to add it.']); } catch (x) {}
  return { ok: true };
}
function twInvSet_(raw) {
  if (String((raw && raw.pin) || '') !== TW_PIN) return { ok: false, error: 'wrong manager code' };
  const f = twInvRead_(), p = f.pending.filter(function (x) { return x.id === raw.pid; })[0];
  if (!p || (raw.status !== 'added' && raw.status !== 'rejected')) return { ok: false, error: 'bad request' };
  if (p.status === 'pending') {
    p.status = raw.status;
    if (raw.status === 'added') { p.code = 'N' + (f.extra.length + 1); f.extra.push({ c: p.code, n: p.name, cat: p.cat }); }
    logWriteJson_(TW_INVF, f);
  }
  return { ok: true };
}

/* ---- daily rollover (12:05 AM Pacific, runs on Google's servers; tablets can be off) ---- */
// Writes ONE record per day into that week's file: every task still open from earlier days, now on today's list as URGENT / OVERDUE.
// Nothing is copied or duplicated: each task keeps its single record (original collaborator, due date, text, creation date).
// Running it again the same day does nothing (the day's record already exists).
function twRollover_(force) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const now = logNow_();
    if (!force && now.min < 5) return { ok: true, skip: 'before 12:05 AM' };
    const cfg = twConfig_(), today = now.date, wk = logMonday_(today);
    if (today < cfg.start) return { ok: true, skip: 'before start date' };
    const d = twReadWeek_(wk, true);
    if (!d.roll) d.roll = {};
    if (d.roll[today]) return { ok: true, dup: true, date: today, n: d.roll[today].n };
    const st = twStats_(cfg, d.comps, twCarry_(cfg, wk), wk, now, false);
    const items = st.items.filter(function (x) { return x.st === 'overdue'; }).map(function (x) {
      return { tid: x.tid, emp: x.emp, orig: x.orig, due: x.due, days: x.late, cr: x.cr, fromWeek: x.carry ? 1 : 0 };
    });
    d.roll[today] = { at: logStamp_(), n: items.length, items: items };
    logWriteJson_(twWeekFile_(wk), d);
    return { ok: true, date: today, n: items.length };
  } finally { try { lock.releaseLock(); } catch (x) {} }
}
// Week files from `from` to `to` (inclusive), oldest first.
function twWeekRange_(cfg, from, to) {
  const out = []; let w = logMonday_(from) < logMonday_(cfg.start) ? logMonday_(cfg.start) : logMonday_(from), n = 0;
  while (w <= to && n++ < 530) { out.push({ week: w, comps: twReadWeek_(w, false).comps }); w = logAddDays_(w, 7); }
  return out;
}
// Manager history: one collaborator's task records (or everyone's with emp empty) for the last N weeks, plus weekly % and recurring problems.
function twHistGet_(emp, weeks) {
  const cfg = twConfig_(), now = logNow_(), cur = logMonday_(now.date), n = Math.max(1, Math.min(52, +weeks || 8));
  const h = twHist_(cfg, twWeekRange_(cfg, logAddDays_(cur, -7 * (n - 1)), cur), now);
  const rec = twRecurring_(cfg, h, 2);
  const it = emp ? h.items.filter(function (x) { return x.emp === emp; }) : [];   // pick a collaborator to download records
  return { ok: true, now: now, weeks: n, emp: emp || '', items: it, summary: h.summary, recurring: rec };
}

/* ---- weekly collaborator performance report (PDF) ---- */
function twStatusTxt_(x) { return x.st === 'done' ? 'Done' : x.st === 'excused' ? 'Excused' : x.st === 'overdue' ? (x.carry ? 'OVERDUE ' + x.late + 'd' : 'MISSING') : x.st === 'pending' ? 'Pending' : 'Scheduled'; }
function twFmtTs_(ts) { const m = /^(\d{4}-\d{2}-\d{2})T(\d{2}):(\d{2})/.exec(ts || ''); if (!m) return ''; const h = +m[2]; return logFmt_(m[1]) + ' ' + ((h % 12) || 12) + ':' + m[3] + (h < 12 ? ' AM' : ' PM'); }
function twPct_(p) { return p === null ? 'n/a' : p + '%'; }
function twPdfHtml_(cfg, st, week, now, hist) {
  const sun = logAddDays_(week, 6), S = st.store, tk = {};
  cfg.tasks.forEach(function (t) { tk[t.id] = t; });
  const th = 'border:1px solid #000;padding:3px 4px;background:#DDD;font-size:9px;text-align:left';
  const td = 'border:1px solid #000;padding:3px 4px;font-size:9px;vertical-align:top';
  const col = function (p) { return p === null || p === undefined ? '#EEE' : p >= 90 ? '#CDEBCB' : p >= 70 ? '#FFF3B0' : '#F4B6B6'; };
  const cell = function (v, ex) { return '<td style="' + td + (ex || '') + '">' + v + '</td>'; };
  const hd = function (a) { return '<tr>' + a.map(function (x) { return '<th style="' + th + '">' + x + '</th>'; }).join('') + '</tr>'; };
  const sumRow = function (n, a) { return '<tr>' + cell(logEsc_(n), ';font-weight:bold') + cell(twPct_(a.pct), ';font-weight:bold;background:' + col(a.pct)) + cell(a.due) + cell(a.done) + cell(a.onTime) + cell(a.lateDone) + cell(a.incomplete) + cell(a.overdue) + cell(a.rolled) + cell(a.excused) + cell(a.prepDone + ' / ' + a.prepDue) + cell(a.assignedDone + ' / ' + a.assigned) + '</tr>'; };
  let h = '<html><body style="font-family:Arial,Helvetica,sans-serif;color:#000">';
  h += '<div style="font-size:18px;font-weight:bold">COLLABORATOR TASK PERFORMANCE &mdash; ' + LOG_NAME + '</div>';
  h += '<div style="font-size:11px;margin:2px 0 8px">Week: Monday ' + logFmt_(week) + ' to Sunday ' + logFmt_(sun) + ' &nbsp;|&nbsp; Semana: lunes a domingo &nbsp;|&nbsp; Generated ' + twFmtTs_(logStamp_()) + '</div>';
  h += '<div style="font-size:14px;font-weight:bold;margin:6px 0">Overall store completion: ' + twPct_(S.pct) + ' (' + S.done + ' of ' + S.due + ' tasks due) &nbsp; On time: ' + S.onTime + ' &nbsp; Late: ' + S.lateDone + ' &nbsp; Incomplete: ' + S.incomplete + '</div>';
  h += '<table style="border-collapse:collapse;width:100%">' + hd(['Collaborator', 'Completed %', 'Assigned (due)', 'Completed', 'On time', 'Late', 'Incomplete', 'Overdue (open)', 'Urgent carryover', 'Excused', 'Prep List done', 'Extra assigned']);
  st.order.forEach(function (n) { h += sumRow(n, st.byEmp[n]); });
  h += sumRow('STORE TOTAL', S) + '</table>';
  h += '<p style="font-size:8px">Completed % = completed &divide; tasks due. Future tasks and tasks excused by a manager are not counted. On time = completed on its due day. Late = completed on a later day. Incomplete = due this week and not completed. Overdue (open) = still not completed after its day, including tasks carried over from earlier weeks. Urgent carryover = rolled over to a later day at least once (still open, or completed late). Completion time = when the collaborator checked it on the tablet.</p>';
  const open = st.items.filter(function (x) { return x.st === 'overdue'; }).sort(function (a, b) { return b.late - a.late || a.emp.localeCompare(b.emp); });
  h += '<div style="font-size:14px;font-weight:bold;margin:10px 0 4px">TASKS REQUIRING MANAGEMENT ATTENTION (' + open.length + ' open overdue)</div>';
  if (open.length) {
    h += '<table style="border-collapse:collapse;width:100%">' + hd(['Collaborator', 'Task', 'Created', 'Original due', 'Days overdue', 'Notes']);
    open.slice(0, 150).forEach(function (x) { const t = tk[x.tid] || {}; h += '<tr style="background:#F4B6B6">' + cell(logEsc_(x.emp), ';font-weight:bold') + cell(logEsc_(t.en || '')) + cell(logFmt_(x.cr)) + cell(logFmt_(x.due)) + cell(x.late, ';font-weight:bold') + cell([x.carry ? 'from an earlier week' : '', x.reas ? 'reassigned by manager' : '', x.pri === 'high' ? 'HIGH priority' : ''].filter(Boolean).join(', ')) + '</tr>'; });
    h += '</table>' + (open.length > 150 ? '<p style="font-size:9px">' + (open.length - 150) + ' more are listed on each collaborator page.</p>' : '');
  } else h += '<p style="font-size:10px">None. Every task due so far was completed or handled.</p>';
  if (hist && hist.recurring && hist.recurring.length) {
    h += '<div style="font-size:12px;font-weight:bold;margin:8px 0 4px">Recurring problems (missed or late 2+ times in the last ' + hist.summary.length + ' weeks)</div><table style="border-collapse:collapse;width:100%">' + hd(['Task', 'Now assigned to', 'Times scheduled', 'Missed', 'Late', 'Still open']);
    hist.recurring.slice(0, 40).forEach(function (r) { const t = tk[r.tid] || {}; h += '<tr>' + cell(logEsc_(t.en || '')) + cell(logEsc_(r.emp || '')) + cell(r.n) + cell(r.miss) + cell(r.late) + cell(r.open) + '</tr>'; });
    h += '</table>';
  }
  const exc = st.items.filter(function (x) { return x.st === 'excused'; });
  if (exc.length) {
    h += '<div style="font-size:12px;font-weight:bold;margin:8px 0 4px">Handled / excused by a manager (' + exc.length + ')</div><table style="border-collapse:collapse;width:100%">' + hd(['Collaborator', 'Task', 'Due', 'Manager note', 'When']);
    exc.forEach(function (x) { h += '<tr>' + cell(logEsc_(x.emp)) + cell(logEsc_((tk[x.tid] || {}).en || '')) + cell(logFmt_(x.due)) + cell(logEsc_(x.ex)) + cell(twFmtTs_(x.exAt)) + '</tr>'; });
    h += '</table>';
  }
  if (hist && hist.summary && hist.summary.length > 1) {
    const ws = hist.summary;
    h += '<div style="font-size:12px;font-weight:bold;margin:10px 0 4px">Week-by-week completion % (compare performance)</div><table style="border-collapse:collapse;width:100%">' + hd(['Collaborator'].concat(ws.map(function (w) { return 'Wk of ' + w.week.slice(5); })));
    st.order.concat(['STORE']).forEach(function (n) {
      h += '<tr>' + cell(logEsc_(n), ';font-weight:bold') + ws.map(function (w) { const a = n === 'STORE' ? w.store : w.byEmp[n], p = a ? a.pct : null; return cell(twPct_(p === undefined ? null : p) + (a && a.due ? ' <span style="color:#555">(' + a.done + '/' + a.due + ')</span>' : ''), ';background:' + col(p)); }).join('') + '</tr>';
    });
    h += '</table>';
  }
  st.order.forEach(function (n) {
    const a = st.byEmp[n], mine = st.items.filter(function (x) { return x.emp === n; }).sort(function (x, y) { return (x.due + (x.carry ? '0' : '1')).localeCompare(y.due + (y.carry ? '0' : '1')); });
    h += '<div style="page-break-before:always"></div><div style="font-size:16px;font-weight:bold">' + logEsc_(n) + ' &mdash; ' + twPct_(a.pct) + ' completed</div>';
    h += '<div style="font-size:10px;margin:3px 0 6px">Assigned (due) ' + a.due + ' | Completed ' + a.done + ' | On time ' + a.onTime + ' | Late ' + a.lateDone + ' | Incomplete ' + a.incomplete + ' | Overdue (open) ' + a.overdue + ' | Urgent carryover ' + a.rolled + ' | Excused ' + a.excused + '</div>';
    h += '<table style="border-collapse:collapse;margin-bottom:6px"><tr>' + a.daily.map(function (d, i) { return '<th style="' + th + '">' + LOG_DOW[i].slice(0, 3) + ' ' + d.date.slice(5) + '</th>'; }).join('') + '</tr><tr>' + a.daily.map(function (d) { return '<td style="' + td + ';background:' + col(d.pct) + '">' + d.done + '/' + (d.done + d.missing) + ' (' + twPct_(d.pct) + ')</td>'; }).join('') + '</tr></table>';
    h += '<table style="border-collapse:collapse;width:100%">' + hd(['Due', 'Task (EN / ES)', 'Frequency', 'Created', 'Status', 'Completed by / when', 'On time?', 'Days overdue', 'Notes']);
    mine.forEach(function (x) {
      const t = tk[x.tid] || {}, bad = x.st === 'overdue';
      const ot = x.st === 'done' ? (x.lateDone ? 'LATE' : 'On time') : x.st === 'excused' ? 'Excused' : '';
      h += '<tr' + (bad ? ' style="background:#F4B6B6"' : x.st === 'done' ? (x.lateDone ? ' style="background:#FFF3B0"' : ' style="background:#E8F5E6"') : '') + '>' + cell(logFmt_(x.due)) + cell(logEsc_(t.en || '') + '<br><i style="color:#555">' + logEsc_(t.es || '') + '</i>') + cell(logEsc_(t.a || '')) + cell(logFmt_(x.cr)) + cell(twStatusTxt_(x), ';font-weight:bold') + cell(x.st === 'done' ? logEsc_(x.by) + ' ' + twFmtTs_(x.cts || x.ts) : '') + cell(ot, x.lateDone ? ';font-weight:bold' : '') + cell(x.daysOver || '') + cell([x.carry ? 'carried over from ' + logFmt_(x.orig) : '', x.moved ? 'moved from ' + logFmt_(x.orig) : '', x.reas ? 'reassigned for this day' : '', x.st === 'excused' ? 'manager: ' + logEsc_(x.ex) : '', x.src === 'assigned' ? 'ASSIGNED' : '', x.link ? 'Prep List' : ''].filter(Boolean).join(', ')) + '</tr>';
    });
    h += '</table>';
  });
  return h + '</body></html>';
}
function twWeekPdf_(week, final) {
  const cfg = twConfig_(), now = logNow_(), d = twReadWeek_(week, false);
  const st = twStats_(cfg, d.comps, twCarry_(cfg, week), week, now, final);
  const hw = twWeekRange_(cfg, logAddDays_(week, -49), week), hist = twHist_(cfg, hw, now);
  hist.recurring = twRecurring_(cfg, hist, 2);
  const name = LOG_NAME + ' - Collaborator Task Performance - ' + week + ' to ' + logAddDays_(week, 6);
  const blob = Utilities.newBlob(twPdfHtml_(cfg, st, week, now, hist), 'text/html', name + '.html').getAs('application/pdf').setName(name + '.pdf');
  return { blob: blob, st: st, name: name, hist: hist };
}
function twPdf_(week) {
  const wk = logMonday_(/^\d{4}-\d{2}-\d{2}$/.test(week || '') ? week : logNow_().date);
  const r = twWeekPdf_(wk, wk < logMonday_(logNow_().date));
  return { ok: true, name: r.name + '.pdf', b64: Utilities.base64Encode(r.blob.getBytes()) };
}
function twSendWeek_(week, test) {
  const r = twWeekPdf_(week, !test), S = r.st.store, sun = logAddDays_(week, 6);
  const fm = function (ds) { return logFmt_(ds).toUpperCase(); };
  const nk = noviTasks_(twConfig_(), r);
  const lines = r.st.order.map(function (n) { const a = r.st.byEmp[n]; return n + ': ' + twPct_(a.pct) + ' (' + a.done + ' of ' + a.due + ' completed: ' + a.onTime + ' on time, ' + a.lateDone + ' late; ' + a.incomplete + ' incomplete, ' + a.overdue + ' overdue, ' + a.rolled + ' urgent carryover)'; });
  const attn = r.st.items.filter(function (x) { return x.st === 'overdue'; }).length;
  const detail = LOG_NAME + ' weekly collaborator task performance\nWeek: Monday ' + logFmt_(week) + ' to Sunday ' + logFmt_(sun) + '\n\n' +
    'Overall store task completion: ' + twPct_(S.pct) + ' (' + S.done + ' of ' + S.due + ' tasks due)\n' +
    'Completed on time: ' + S.onTime + ' | Completed late: ' + S.lateDone + ' | Incomplete: ' + S.incomplete + '\n' +
    'Overdue (still open): ' + S.overdue + ' | Urgent carryover: ' + S.rolled + ' | Excused by manager: ' + S.excused + '\n' +
    'Tasks requiring management attention: ' + attn + (r.hist.recurring.length ? ' | Recurring problem tasks: ' + r.hist.recurring.length : '') + '\n' +
    'Weekly Prep List tasks completed: ' + S.prepDone + ' of ' + S.prepDue + '\nAdditional assigned tasks completed: ' + S.assignedDone + ' of ' + S.assigned + '\n\n' + lines.join('\n') +
    '\n\nResumen en español: finalización general de tareas ' + twPct_(S.pct) + '. El reporte completo por empleado está en el PDF adjunto.\nThe full report for each collaborator is attached as a PDF.';
  const res = { t: Date.now(), at: logStamp_(), ok: false, msg: '' };
  try {
    noviMail_({ to: LOG_TO, test: test, report: 'Weekly Collaborator Task Report', period: 'the week of Monday ' + logFmt_(week) + ' to Sunday ' + logFmt_(sun), date: noviRange_(week, sun),
      summary: nk.summary, obs: nk.obs, recs: nk.recs, facts: nk.facts, detail: detail, attachments: [r.blob] });
    res.ok = true; res.msg = 'Accepted by Gmail/MailApp with no error. Daily send quota left: ' + MailApp.getRemainingDailyQuota();
  } catch (err) { res.msg = 'FAILED: ' + String(err); }
  if (!test) {
    const d = twReadWeek_(week, true);
    d.report.attempts.push(res);
    if (res.ok) {
      d.report.status = 'SENT'; d.report.sentAt = res.at;
      // Frozen copy of the numbers as sent, so later changes never rewrite an old week's result.
      const be = {}; r.st.order.forEach(function (n) { const a = r.st.byEmp[n]; be[n] = { pct: a.pct, due: a.due, done: a.done, onTime: a.onTime, lateDone: a.lateDone, incomplete: a.incomplete, overdue: a.overdue, rolled: a.rolled, excused: a.excused }; });
      d.report.snap = { at: res.at, store: { pct: S.pct, due: S.due, done: S.done, onTime: S.onTime, lateDone: S.lateDone, incomplete: S.incomplete, overdue: S.overdue, rolled: S.rolled, excused: S.excused }, byEmp: be };
    }
    logWriteJson_(twWeekFile_(week), d);
  }
  return res;
}
// Called by the SAME two triggers as the temperature log (no extra schedule). Sends each finished week once; retries a failure after 25 minutes.
function twProcess_() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const cfg = twConfig_(), now = logNow_(), cur = logMonday_(now.date);
    let w = logMonday_(cfg.start), n = 0;
    while (w <= cur && n++ < 520) {
      const d = twReadWeek_(w, true);
      const sun = logAddDays_(w, 6), due = now.date > sun || (now.date === sun && now.min >= 1380);   // Sunday 11:00 PM
      if (due && d.report.status !== 'SENT') {
        const last = d.report.attempts[d.report.attempts.length - 1];
        if (!(last && Date.now() - last.t < 25 * 60000)) twSendWeek_(w, false);
      }
      w = logAddDays_(w, 7);
    }
  } finally { try { lock.releaseLock(); } catch (x) {} }
}
// Run from the editor to test the collaborator report email (current week, labelled TEST, records nothing).
function sendTestTaskReport() { Logger.log(JSON.stringify(twSendWeek_(logMonday_(logNow_().date), true))); }

/* ---- shared Prep list + Close-out checklist (all tablets see the same one) ---- */
const PREP_FILE = 'prep-state.json';
function prepLoad_() {
  const d = logReadJson_(PREP_FILE) || {};
  if (!d.l) d.l = {}; if (!d.pw) d.pw = {}; if (!d.pwt) d.pwt = {};
  return d;
}
function prepDayStart_(ds) { try { return Utilities.parseDate(ds, LOG_TZ, 'yyyy-MM-dd').getTime(); } catch (e) { return Date.now(); } }
// Rolls every location to today if needed. Returns true if anything changed.
function prepRollDoc_(doc, today, ensureLoc) {
  let ch = false;
  if (ensureLoc && !doc.l[ensureLoc]) { doc.l[ensureLoc] = { d: '', n: {}, x: {}, u: {}, c: [], o: {}, b: {}, t: {}, rt: 0 }; ch = true; }
  Object.keys(doc.l).forEach(function (loc) {
    const st = doc.l[loc];
    if (st.d !== today) {
      pRoll_(st, doc.pw, today); st.t = {}; st.rt = prepDayStart_(today); ch = true;
      const lim = logAddDays_(today, -35);
      Object.keys(doc.pwt).forEach(function (k) { if (k.slice(2, 12) < lim) delete doc.pwt[k]; });
    }
  });
  return ch;
}
function prepSync_(b) {
  const loc = String(b.loc || '').slice(0, 50);
  if (!loc) return { ok: false, error: 'location required' };
  const pcfg = pcfgLoad_();   // PREP LIST / WEEK VIEW from the Google Sheet, if set up
  const doc = prepLoad_(), today = logNow_().date;
  let dirty = prepRollDoc_(doc, today, loc);
  const st = doc.l[loc], nowMs = Date.now(), chs = Array.isArray(b.ch) ? b.ch.slice(0, 600) : [];
  chs.forEach(function (c) {
    const k = String(c && c.k || '');
    if (k.length < 3 || k.length > 120 || k.charAt(1) !== ':' || 'nxucowbe'.indexOf(k.charAt(0)) < 0) return;
    const isW = k.charAt(0) === 'w';
    if (isW && !/^w:\d{4}-\d{2}-\d{2}\|/.test(k)) return;
    let v = c.v;
    if (k.charAt(0) === 'c') { v = (v && typeof v === 'object' && v.n) ? { n: String(v.n), g: String(v.g || '') } : 0; } else if (k.charAt(0) === 'b') { v = ['AM', 'MID', 'PM'].indexOf(v) >= 0 ? v : 0; } else if (k.charAt(0) === 'e') { v = (typeof v === 'string' && /^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ .'-]{0,29}$/.test(v)) ? v : 0; } else { v = v ? 1 : 0; }
    const t = Math.min(Number(c.t) || 0, nowMs + 60000), tm = isW ? doc.pwt : st.t;
    if (c.f) { if (!v || tm[k] !== undefined) return; }               // first connection of a device: only ADDS things, never undoes
    else {
      if (!isW && t < (st.rt || 0)) return;                            // made before today started: old, ignore
      if (t < (tm[k] || 0)) return;                                    // somebody changed it more recently
      tm[k] = t;
    }
    pSet_(st, doc.pw, k, v); dirty = true;
  });
  if (dirty) logWriteJson_(PREP_FILE, doc);
  return { ok: true, d: st.d, n: st.n, x: st.x, u: st.u, c: st.c, o: st.o, b: st.b || {}, e: st.e || {}, pw: doc.pw, now: nowMs, pc: (pcfg && b.cv !== pcfg.v) ? pcfg : undefined };
}
// Called by the 12:05 AM trigger and the 30-minute trigger so the new day is ready even if no tablet is on.
function prepRoll_() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const now = logNow_();
    if (now.min < 5) return;
    pcfgLoad_();
    const doc = prepLoad_();
    if (prepRollDoc_(doc, now.date, '')) logWriteJson_(PREP_FILE, doc);
  } finally { try { lock.releaseLock(); } catch (x) {} }
}

/* ---- edit a past temperature reading (keeps the old values in the entry's edit history) ---- */
function logEdit_(raw) {
  raw = raw || {};
  const c = logConfig_(), now = logNow_();
  const bad = function (m) { return { ok: false, fatal: true, error: 'invalid: ' + m }; };
  const ref = String(raw.ref || ''), date = String(raw.date || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(ref)) return bad('id');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return bad('date');
  if (date > logAddDays_(now.date, 1)) return bad('date in the future');
  const wk = logMonday_(date), d = logReadWeek_(wk, true);
  const e = d.entries.filter(function (x) { return x.id === ref; })[0];
  if (!e) return bad('reading not found');
  const empty = raw.empty === true, temp = empty ? null : Number(raw.temp);
  if (!empty && (!isFinite(temp) || temp < 0 || temp > 400 || raw.temp === '' || raw.temp === null)) return bad('temperature');
  const result = empty ? 'EMPTY' : temp < LOG_FAIL_BELOW ? 'FAIL' : 'PASS';
  const action = String(raw.action || '').trim().slice(0, 500);
  if (result === 'FAIL' && !action) return bad('corrective action required below ' + LOG_FAIL_BELOW + 'F');
  const item = c.items.filter(function (x) { return x.equip === e.equip && x.food === e.food; })[0];
  const name = (item && item.flex || e.name) && !empty ? String(raw.name || e.name || '').trim().slice(0, 40) : '';
  if (!e.edits) e.edits = [];
  e.edits.push({ at: logStamp_(), temp: e.temp, empty: !!e.empty, result: e.result, action: e.action || '', name: e.name || '', disd: !!e.disd });
  e.temp = empty ? null : Math.round(temp * 10) / 10; e.empty = empty; e.result = result; e.action = action; e.name = name;
  e.disd = !empty && raw.disd === true; e.sts = logStamp_();
  logWriteWeek_(wk, d);
  return { ok: true, entry: e };
}

/* ---- alarms: on/off switches and extra alarms, shared by all tablets, changed only with the manager code ---- */
function alGet_() {
  const d = logReadJson_('alarms.json') || {};
  return { ok: true, t: Number(d.t) || 0, off: d.off || {}, custom: d.custom || [], list: d.list || null };
}
function alClean_(a) {
  if (!a || !/^[A-Za-z0-9_-]{1,20}$/.test(String(a.id)) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(String(a.t)) || !String(a.n || '').trim() || !Array.isArray(a.d)) return null;
  const d = a.d.map(Number).filter(function (x, i, r) { return x >= 0 && x <= 6 && r.indexOf(x) === i; });
  if (!d.length) return null;
  return { id: String(a.id), t: String(a.t), n: String(a.n).trim().slice(0, 40).toUpperCase(), d: d, go: ['close', 'temps', 'prep'].indexOf(a.go) >= 0 ? a.go : '', tk: alTk_(a), who: String(a.who || '').trim().toUpperCase().slice(0, 40) };
}
// From a tablet in Manager mode. Also rewrites the ALARMS tab so the Google Sheet always matches.
function alSet_(b) {
  if (String((b && b.pin) || '') !== TW_PIN) return { ok: false, error: 'wrong manager code' };
  const cur = alGet_(), t = Math.min(Number(b.t) || 0, Date.now() + 60000);
  if (t < cur.t) return cur;
  const off = {}; Object.keys(b.off || {}).slice(0, 100).forEach(function (k) { if (b.off[k] && /^[A-Za-z0-9_-]{1,20}$/.test(k)) off[k] = 1; });
  const list = Array.isArray(b.list) ? b.list.slice(0, 80).map(alClean_).filter(Boolean)
    : alList_({ custom: (Array.isArray(b.custom) ? b.custom : []).slice(0, 40).map(alClean_).filter(Boolean) });
  if (!list.length) return { ok: false, error: 'no alarms' };
  logWriteJson_('alarms.json', { t: t, list: list, off: off, from: 'tablet' });
  try { const sh = tabFind_(SpreadsheetApp.getActiveSpreadsheet(), AL_TAB); if (sh) { alSheetWrite_(sh, list, off, 'Changed on a tablet ' + twFmtTs_(logStamp_())); tabLook_(sh); } } catch (x) {}
  return { ok: true, t: t, off: off, custom: [], list: list };
}

/* ---- run ONCE from the editor: removes the old unused tabs (Sheet1, Counts) from the Google Sheet ---- */
// Tabs the system uses are never touched: Stock, Reports, Close outs, Orders, Dashboard, Purchase list.
function cleanupSheetTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), keep = Object.keys(TAB_NAMES).map(tabName_);
  const log = [];
  ss.getSheets().forEach(function (sh) {
    const n = sh.getName();
    if (keep.indexOf(n) >= 0) return;
    const rows = sh.getLastRow(), old = (n === 'Sheet1' || n === 'Counts');
    if (old && ss.getSheets().length > 1 && (n === 'Sheet1' ? rows <= 1 : !!tabFind_(ss, PL))) { ss.deleteSheet(sh); log.push('Deleted: ' + n + ' (' + rows + ' rows)'); }
    else log.push('Left alone: ' + n + ' (' + rows + ' rows)' + (n === 'Sheet1' ? ' - it has data, delete it yourself if you do not need it' : ''));
  });
  log.push('Tabs now: ' + ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
  Logger.log(log.join('\n'));
  return log;
}

/* ===================== COLLABORATOR TASKS FROM THE GOOGLE SHEET ===================== */
// The "Tasks" and "Collaborators" tabs are the ONE place to change collaborator tasks. Tablets only check tasks off.
// A change applies from the day it is made, so earlier days keep who was assigned and what was due.
const TW_MGR_EDITS = false;   // false = managers cannot change tasks from a tablet (the Google Sheet is used instead)
const TS_TAB = 'Tasks', TC_TAB = 'Collaborators';
const TS_HEAD = ['ID (do not change)', 'Collaborator', 'Task', 'How often', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN', 'Start date', 'Active', 'Prep list item', 'Status (filled in automatically)'];
const TS_SPARE = 20;
// Days are checkboxes (MON..SUN, columns E..K) like the ALARMS and PREP LIST tabs. Older tabs had one "Days" text column: converted once.
function tsDaysCols_(sh) {
  try {
    if (String(sh.getRange(1, 5).getValue() || '').trim() !== 'Days') return false;
    const n = Math.max(0, sh.getLastRow() - 1), old = n ? sh.getRange(2, 4, n, 2).getValues() : [];
    for (let i = 0; i < 6; i++) sh.insertColumnAfter(5);
    sh.getRange(1, 5, 1, 7).setValues([TS_DAYS.slice()]).setFontWeight('bold');
    const vals = old.map(function (r) {
      const how = String(r[0] || '').trim().toLowerCase(), d = [];
      tsPlain_(r[1]).split(/[\s,\/;&]+|\bAND\b|\bY\b/).forEach(function (w) { const s3 = w.slice(0, 3); let k = TS_DAYS.indexOf(s3); if (k < 0) k = TS_DAYS_ES.indexOf(s3); if (k >= 0 && d.indexOf(k) < 0) d.push(k); });
      const all = /^(every ?day|daily|diario|todos los d)/.test(how);
      return TS_DAYS.map(function (x, k) { return all || d.indexOf(k) >= 0; });
    });
    if (n) sh.getRange(2, 5, n, 7).setValues(vals);
    tsFormat_(sh, n);
    return true;
  } catch (x) { return false; }
}
function tsFormat_(sh, n) {
  const rows = Math.max(1, n + TS_SPARE);
  try { sh.getRange(2, 5, rows, 7).insertCheckboxes(); } catch (x) {}
  try {
    sh.setFrozenRows(1); [110, 190, 420, 150, 52, 52, 52, 52, 52, 52, 52, 110, 70, 170, 380].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange(1, 5, rows + 1, 7).setHorizontalAlignment('center');
    sh.getRange(2, 12, rows, 1).setNumberFormat('m/d/yyyy');
    sh.getRange(2, 4, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TS_HOW, true).setAllowInvalid(true).build());
    sh.getRange(2, 13, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
  } catch (x) {}
}
// A row of the tab (A..N) -> the parser's row shape [id, collaborator, task, '', how often, days text, start, active, prep item]
function tsRowIn_(r) {
  const ticks = r.slice(4, 11).map(alOnCell_), on = ticks.filter(Boolean).length;
  let how = String(r[3] === null ? '' : r[3]).trim(), days = '';
  const low = how.toLowerCase();
  if (/^(every ?day|daily|diario|todos los d)/.test(low)) { if (on > 0 && on < 7) { how = 'Weekly'; days = TS_DAYS.filter(function (x, k) { return ticks[k]; }).join(', '); } }
  else if (/week|semana|fixed|days|prep/.test(low)) days = TS_DAYS.filter(function (x, k) { return ticks[k]; }).join(', ');
  return [r[0], r[1], r[2], '', how, days, r[11], r[12], r[13]];
}
// Tasks are typed in English only. The Spanish (ES button on the tablets) is made automatically:
// Google Translate now; once the ZELYNA_AI_KEY Script property is added, the AI translates (kitchen words kept) and redoes the Google ones.
const TR_FILE = 'translations.json';
function trLoad_() { let d = null; try { d = logReadJson_(TR_FILE); } catch (x) {} return d && d.es ? d : { es: {} }; }
function trKey_() { return zelKey_(); }
// list of English texts -> {en: {es, src:'ai'|'g'}}
function trTranslate_(list) {
  const out = {}; list = list.filter(function (x, i) { return x && list.indexOf(x) === i; }).slice(0, 60);
  if (!list.length) return out;
  const key = trKey_();
  if (key) {
    try {
      const sys = 'You translate short restaurant tasks for the kitchen staff of ' + NOVI_LOC + ', a Mexican restaurant (quesabirrias, tacos). Translate each English task to natural Mexican Spanish as used in a taqueria kitchen. ' +
        'Keep Spanish kitchen words that are already Spanish (PICAR, ASAR, MARINAR, CONSOMÉ, BIRRIA, ASADA, AL PASTOR, CEBOLLA...). Keep names, numbers and units. Write in CAPITAL LETTERS. Reply with JSON only: {"English text": "SPANISH TEXT", ...}';
      const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: NOVI_MODEL, max_tokens: 4000, system: sys, messages: [{ role: 'user', content: JSON.stringify(list) }] }) });
      if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
      const j = JSON.parse(res.getContentText()), txt = (j.content || []).map(function (c) { return c.text || ''; }).join(''), m = /\{[\s\S]*\}/.exec(txt), o = JSON.parse(m[0]);
      list.forEach(function (en) { const v = String(o[en] || '').trim(); if (v) out[en] = { es: v.toUpperCase().slice(0, 200), src: 'ai' }; });
    } catch (e) {}
  }
  list.forEach(function (en) {
    if (out[en]) return;
    try { const v = String(LanguageApp.translate(en, 'en', 'es') || '').trim(); if (v) out[en] = { es: v.toUpperCase().slice(0, 200), src: 'g' }; } catch (x) {}
  });
  return out;
}
// Old tabs had a "Task (Spanish)" column D: its Spanish is kept (as the translation of each task), then the column is removed.
function tsDropEs_(sh) {
  try {
    if (String(sh.getRange(1, 4).getValue() || '').trim() !== 'Task (Spanish)') return false;
    const n = sh.getLastRow() - 1, d = trLoad_();
    if (n > 0) sh.getRange(2, 3, n, 2).getValues().forEach(function (r) { const en = String(r[0] || '').trim().slice(0, 200), es = String(r[1] || '').trim().slice(0, 200); if (en && es && es !== en) d.es[en] = { es: es, src: 'sheet' }; });
    d.at = logStamp_(); logWriteJson_(TR_FILE, d);
    sh.deleteColumn(4); sh.getRange(1, 3).setValue('Task');
    return true;
  } catch (x) { return false; }
}
const TC_HEAD = ['Name', 'Role', 'Shift', 'Active', 'Status (filled in automatically)', 'Spanish role ending (o / a)'];
const TS_DAYS = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'], TS_DAYS_ES = ['LUN', 'MAR', 'MIE', 'JUE', 'VIE', 'SAB', 'DOM'];
const TS_HOW = ['Every day', 'Every 2 days', 'Every 3 days', 'Every other day', 'Weekly', 'Every 2 weeks', 'When on prep list', 'As needed', 'One time'];
const TS_ROLE = { ALICIA: 'Cashier', JENNIFER: 'Cashier', BRISA: 'Cashier', HUGO: 'Cook', ERIKA: 'Cook', MILTZY: 'Cook' };
const TS_SHIFT = { ALICIA: 'AM', BRISA: 'PM', JENNIFER: 'AM and PM', HUGO: 'AM', MILTZY: 'PM' };

function tsKind_(f) { return (f === 'W1' || f === 'W2' || f === 'W3' || f === 'DAYS' || f === 'PREP') ? 'DAYS' : (f === 'E3' || f === 'EV') ? 'EV' : f; }
function tsDaysTxt_(d) { return (d || []).slice().sort().map(function (i) { const s = TS_DAYS[i]; return s.charAt(0) + s.slice(1).toLowerCase(); }).join(', '); }
function tsHowTxt_(t) {
  const k = tsKind_(t.f);
  if (t.f === 'PREP' && !(t.d || []).length) return 'When on prep list';
  return k === 'DAY' ? 'Every day' : k === 'ODD' ? 'Every other day' : k === 'EV' ? 'Every ' + (t.ev || 3) + ' days' : k === 'DAYS' ? 'Weekly' : k === 'BIW' ? 'Every 2 weeks' : k === 'NEED' ? 'As needed' : k === 'ONE' ? 'One time' : 'Not set';
}
function tsRowOf_(t) {
  const k = tsKind_(t.f);
  return [t.id, t.rot ? t.rot.join(', ') : t.emp, t.en || '', tsHowTxt_(t)].concat(TS_DAYS.map(function (x, i) { return k === 'DAY' || ((k === 'DAYS' || k === 'BIW') && (t.d || []).indexOf(i) >= 0); })).concat([
    k === 'ONE' ? (t.date || '') : (k === 'EV' || t.rot) ? (t.an || '') : '', 'Yes', (t.l || []).join(', '), '']);
}
function tsDate_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? null : Utilities.formatDate(v, LOG_TZ, 'yyyy-MM-dd');
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s);
  if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return y + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2); }
  return null;
}
function tsPlain_(s) { return String(s || '').toUpperCase().replace(/[ÁÀ]/g, 'A').replace(/[ÉÈ]/g, 'E').replace(/[ÍÌ]/g, 'I').replace(/[ÓÒ]/g, 'O').replace(/[ÚÙ]/g, 'U'); }
// One row of the Tasks tab -> what it asks for (or the problems found in it).
function tsParseRow_(r, names, today) {
  const p = { err: [], id: String(r[0] || '').trim() };
  p.emps = String(r[1] || '').toUpperCase().split(/[,\/;]+/).map(function (x) { return x.trim(); }).filter(Boolean);
  if (!p.emps.length) p.err.push('Collaborator is empty');
  p.emps.forEach(function (n) { if (names.indexOf(n) < 0 && !twAnyPos_(n)) p.err.push('Unknown collaborator "' + n + '" (add the name on the Collaborators tab first)'); });
  if (p.emps.length > 1 && p.emps.some(twAnyPos_)) p.err.push('Use ANY CASHIER / ANY COOK / ANY DISHES alone (no other names on the same row)');
  p.link = String(r[8] || '').split(',').map(function (x) { return x.trim().toUpperCase(); }).filter(Boolean).slice(0, 5);
  p.en = String(r[2] || '').trim().slice(0, 200); p.es = String(r[3] || '').trim().slice(0, 200) || p.en;
  if (!p.en) p.err.push('Task is empty');
  const how = String(r[4] || '').trim().toLowerCase(); let m;
  if (/^(every ?day|daily|diario|todos los d)/.test(how)) p.f = 'DAY';
  else if (/every other day|alternate|d[ií]a de por medio/.test(how)) p.f = 'ODD';
  else if ((m = /(?:every|cada)\s*(\d+)\s*(?:days?|d[ií]as?)/.exec(how))) { const n = +m[1]; if (n === 1) p.f = 'DAY'; else if (n >= 2 && n <= 30) { p.f = 'EV'; p.ev = n; } else p.err.push('Every N days must be 2 to 30'); }
  else if (/every\s*(2|two|other)\s*weeks?|biweekly|cada\s*2\s*semanas/.test(how)) p.f = 'BIW';
  else if (/prep/.test(how)) p.f = 'PREP';
  else if (/week|semana|fixed|days/.test(how)) p.f = 'DAYS';
  else if (/as needed|seg[uú]n|when needed/.test(how)) p.f = 'NEED';
  else if (/one ?time|once|una vez/.test(how)) p.f = 'ONE';
  else p.err.push('"How often" not understood. Use: ' + TS_HOW.join(', '));
  p.d = [];
  if (p.f === 'DAYS' || p.f === 'BIW' || p.f === 'PREP') {
    tsPlain_(r[5]).split(/[\s,\/;&]+|\bAND\b|\bY\b/).forEach(function (w) { const s = w.slice(0, 3); let i = TS_DAYS.indexOf(s); if (i < 0) i = TS_DAYS_ES.indexOf(s); if (i >= 0 && p.d.indexOf(i) < 0) p.d.push(i); });
    p.d.sort(); if (!p.d.length && p.f !== 'PREP') p.err.push('Pick the day(s) in the Days column, for example: Mon, Thu');
    if (p.f === 'PREP' && !p.link.length) p.err.push('Fill in the Prep list item column');
  }
  const sd = tsDate_(r[6]); if (sd === null) p.err.push('Start date not understood (use 11/3/2026)'); p.start = sd || '';
  if (p.f === 'ONE' && !p.start) p.err.push('One time tasks need a Start date');
  p.active = !/^(no|n|false|0|off|inactive|inactivo)$/i.test(String(r[7] === false ? 'no' : r[7] || '').trim());
  return p;
}
function tsKey_(s) { const k = tsKind_(s.f); return JSON.stringify([k, (k === 'DAYS' || k === 'BIW') ? (s.d || []).slice().sort() : [], k === 'EV' ? (s.ev || 3) : 0, (k === 'EV' || (s.rot && s.rot.length)) ? (s.an || '') : '', k === 'ONE' ? (s.date || '') : '', s.rot && s.rot.length ? s.rot.join(',') : '']); }

// Reads the Tasks and Collaborators tabs and updates the tablets. Runs on every edit of those tabs, every 30 minutes, and from the Birriaz menu.
function twSheetSync_() {
  let ss; try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { return { ok: false, error: 'no spreadsheet' }; }
  const sh = ss && tabFind_(ss, TS_TAB);
  if (!sh) { let pc = null, iv = null; try { pc = pcfgSync_(ss); } catch (x) {} try { iv = invSync_(ss); } catch (x) {} try { alSync_(ss); } catch (x) {} return { ok: false, error: 'no Tasks tab yet (run setupSheets once)', prep: pc, inv: iv }; }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    try { pl2MigrateIfOld_(ss, []); } catch (x) {}
    try { pl2StageMerge_(ss, []); } catch (x) {}
    try { pl2DropShelf_(ss, []); } catch (x) {}
    try { layout3_(ss, []); } catch (x) {}
    try { l3WhoDv_(ss); } catch (x) {}
    try { const P = PropertiesService.getScriptProperties(); if (P.getProperty('DV_LISTS') !== '3') { dvLists_(ss, null); P.setProperty('DV_LISTS', '3'); }
      if (P.getProperty('HIDE_ST') !== '1') { hideStatus_(ss, null); P.setProperty('HIDE_ST', '1'); } } catch (x) {}
    const cfg = twConfig_(), today = logNow_().date, stamp = twFmtTs_(logStamp_()), ops = [];
    // 1. collaborators (old list layout or the W weekly schedule layout)
    const csh = tabFind_(ss, TC_TAB);
    if (csh) {
      const T = tcRead_(csh), st = [];
      T.rows.forEach(function (r) {
        const n = String(r.name || '').trim().toUpperCase();
        if (!n) { st.push(['']); return; }
        if (!/^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ .'-]{0,29}$/.test(n) || twAnyPos_(n) || n === 'ANYONE') { st.push(['⚠ Use letters only (30 max)']); return; }
        const e = cfg.employees.filter(function (x) { return x.name === n; })[0], off = !/^(yes|y|si|sí|true|1|active|activo|)$/i.test(String(r.act === true ? 'yes' : r.act || '').trim()) ? 1 : 0;
        const role = String(r.role || '').trim().slice(0, 20), shift = String(r.shift || '').trim().slice(0, 20), rend = /^\s*[oOaA]\s*$/.test(String(r.rend || '')) ? String(r.rend).trim().toLowerCase() : '';
        const ch = [];
        if (!e) { ops.push({ k: 'addEmp', name: n, from: today }); ops.push({ k: 'empInfo', emp: n, role: role, shift: shift, rend: rend, off: off, from: today }); ch.push('added'); }
        else if ((e.role || '') !== role || (e.shift || '') !== shift || (e.rend || '') !== rend || (e.off ? 1 : 0) !== off) { ops.push({ k: 'empInfo', emp: n, role: role, shift: shift, rend: rend, off: off, from: today }); ch.push('updated'); }
        if (r.wk !== undefined && JSON.stringify((e && e.wk) || null) !== JSON.stringify(r.wk)) { ops.push({ k: 'wk', emp: n, wk: r.wk, from: today }); ch.push('schedule'); }
        const pos = twPosOf_(role), bad = r.bad && r.bad.length ? ' · ⚠ ' + r.bad.join('; ') : '';
        st.push([(ch.length ? (ch[0] === 'added' ? 'Added ' : 'Updated ') + stamp + (ch.indexOf('schedule') >= 0 ? ' (schedule)' : '') : '✓ OK') + (role && !pos ? ' · position not Cashier, Cook or Dishes: no ANY tasks' : '') + bad]);
      });
      tcStWrite_(csh, T, st);
    }
    ops.forEach(function (o) { twApply_(cfg, o); });
    const names = cfg.employees.map(function (e) { return e.name; }), tops = [];
    // 2. tasks
    tsDropEs_(sh); tsDaysCols_(sh);
    const L3 = ts3Is_(sh), n = sh.getLastRow() - 1, rows = (n > 0 ? sh.getRange(2, 1, n, L3 ? TS3_HEAD.length : 14).getValues() : []).map(L3 ? ts3RowIn_ : tsRowIn_), status = [], seen = {};
    const idCol = L3 ? TS3.id + 1 : 1, stCol = L3 ? TS3.st + 1 : 15;
    let added = 0, changed = 0, bad = 0;
    // Spanish: kept while the English stays the same; new or changed English is translated (all at once)
    const tr = trLoad_(), ai = !!trKey_(), need = [];
    rows.forEach(function (r) { const en = String(r[2] || '').trim().slice(0, 200), id = String(r[0] || '').trim(), t0 = id ? twFind_(cfg, id) : null; if (!en) return;
      const c = tr.es[en];
      if (c && c.src === 'g' && ai) need.push(en);                       // the AI redoes Google translations once the key is added
      else if (!c && !(t0 && t0.en === en && t0.es && t0.es !== en)) need.push(en); });
    if (need.length) { const got = trTranslate_(need); Object.keys(got).forEach(function (en) { tr.es[en] = got[en]; }); if (Object.keys(got).length) { tr.at = logStamp_(); try { logWriteJson_(TR_FILE, tr); } catch (x) {} } }
    const esOf = function (en, t0) { const c = tr.es[en]; if (c) return c.es; return (t0 && t0.en === en && t0.es) ? t0.es : en; };
    rows.forEach(function (r, i) {
      if (r.every(function (v) { return String(v === null ? '' : v).trim() === ''; })) { status.push(['']); return; }
      const p = tsParseRow_(r, names, today);
      p.es = esOf(p.en, p.id ? twFind_(cfg, p.id) : null);
      if (p.id && seen[p.id] && L3) { p.id = ''; try { sh.getRange(i + 2, idCol).setValue(''); } catch (x) {} }   // a copied row (the ID column is hidden): it becomes a new task
      if (p.id && seen[p.id]) { status.push(['⚠ This ID is used twice. Clear the ID cell on the new row.']); bad++; return; }
      if (p.id) seen[p.id] = 1;
      if (p.err.length) { status.push(['⚠ ' + p.err.join('; ') + ' (no changes made for this row)']); bad++; return; }
      const t = p.id ? twFind_(cfg, p.id) : null;
      if (p.id && !t && L3) { p.id = ''; try { sh.getRange(i + 2, idCol).setValue(''); } catch (x) {} }   // a row with an ID from elsewhere: added as a new task
      if (p.id && !t) { status.push(['⚠ Unknown ID. Clear the ID cell to add this as a new task.']); bad++; return; }
      const rot = p.emps.length > 1 ? p.emps : null;
      const want = { f: p.f, d: p.d, ev: p.ev || 0, an: '', date: p.f === 'ONE' ? p.start : '', rot: rot };
      if (p.f === 'EV' || rot) want.an = p.start || (t && t.an && (tsKind_(t.f) === 'EV' || t.rot) ? t.an : today);
      if (!t) {
        const id = 's' + Date.now().toString(36) + i;
        tops.push({ k: 'addSheet', task: { id: id, en: p.en, es: p.es, emp: p.emps[0], rot: rot, f: p.f, d: p.d, ev: want.ev, an: want.an, date: want.date, l: p.link, start: today } });
        if (!p.active) tops.push({ k: 'off', tid: id, off: 1, from: today });
        try { sh.getRange(i + 2, idCol).setValue(id); } catch (x) {}
        seen[id] = 1; status.push(['Added ' + stamp + '. It shows on the tablets from today.']); added++; return;
      }
      const ch = [];
      if (!p.active && !t.off) { tops.push({ k: 'off', tid: t.id, off: 1, from: today }); ch.push('turned off'); }
      if (p.active && t.off) { tops.push({ k: 'off', tid: t.id, off: 0, from: today }); ch.push('turned on'); }
      if (!rot && (t.rot || t.emp !== p.emps[0])) { tops.push({ k: 'assign', tid: t.id, emp: p.emps[0], from: today }); ch.push('now ' + p.emps[0]); }
      if (p.en !== t.en || p.es !== t.es) { tops.push({ k: 'clar', tid: t.id, en: p.en, es: p.es }); ch.push('text'); }
      if (tsKey_(want) !== tsKey_(t)) {
        if (want.f === 'DAYS' && p.link.length) want.f = 'PREP';
        tops.push(Object.assign({ k: 'sched', tid: t.id, from: today }, want)); ch.push('schedule');
      }
      if ((t.l || []).join('|') !== p.link.join('|')) { tops.push({ k: 'link', tid: t.id, l: p.link }); ch.push('prep list item'); }
      if (ch.length) changed++;
      status.push([ch.length ? 'Updated ' + stamp + ' (' + ch.join(', ') + '). Applies from today.' : '✓ OK']);
    });
    // 3. rows that were deleted from the tab: turned off from today (guard against an accidentally emptied tab)
    const gone0 = cfg.tasks.filter(function (t) { return !seen[t.id] && !t.off && !(t.f === 'ONE' && t.date < today); });
    // rows moved to the PREP LIST by "Split rows by TYPE" are turned off for sure (the owner confirmed them)
    const okIds = {}; try { String(PropertiesService.getScriptProperties().getProperty('TS_GONE_OK') || '').split(',').forEach(function (x) { if (x) okIds[x] = 1; }); } catch (x) {}
    const forced = gone0.filter(function (t) { return okIds[t.id]; }), gone = gone0.filter(function (t) { return !okIds[t.id]; });
    let skipped = 0;
    if (gone.length > 10 || !Object.keys(seen).length) skipped = gone.length;
    else gone.forEach(function (t) { tops.push({ k: 'off', tid: t.id, off: 1, from: today }); });
    forced.forEach(function (t) { tops.push({ k: 'off', tid: t.id, off: 1, from: today }); });
    if (forced.length || Object.keys(okIds).length) { try { PropertiesService.getScriptProperties().deleteProperty('TS_GONE_OK'); } catch (x) {} }
    tops.forEach(function (o) { twApply_(cfg, o); });
    const all = ops.concat(tops);
    if (all.length) {
      if (!cfg.log) cfg.log = [];
      all.forEach(function (o) { cfg.log.push({ id: 'sheet', k: o.k, tid: o.tid || (o.task && o.task.id) || '', orig: o.from || '', emp: o.emp || o.name || '', note: 'Google Sheet', at: logStamp_() }); });
      cfg.sheetAt = logStamp_();
      logWriteJson_(TW_CFG, cfg);
    }
    try {
      stWrite_(sh, 2, stCol, status);
      const hd = 'Status (filled in automatically)' + (skipped ? ' - ⚠ ' + skipped + ' tasks are missing from this tab, so none were turned off. To remove a task set Active to No.' : '');
      if (String(sh.getRange(1, stCol).getValue() || '') !== hd) sh.getRange(1, stCol).setValue(hd);
    } catch (x) {}
    let pc = null, iv = null; try { pc = pcfgSync_(ss); } catch (x) { pc = { ok: false, error: String(x) }; }
    try { iv = invSync_(ss); } catch (x) { iv = { ok: false, error: String(x) }; }
    let al = null; try { al = alSync_(ss); } catch (x) { al = { ok: false, error: String(x) }; }
    return { ok: true, changes: all.length, added: added, changed: changed, problems: bad, removed: skipped ? 0 : gone.length + forced.length, skipped: skipped, prep: pc, inv: iv, alarms: al };
  } finally { try { lock.releaseLock(); } catch (x) {} }
}
// Installed by setupTaskSheet: runs when someone edits the Tasks or Collaborators tab.
// Picking "Every day" in the How often column of the COLLABORATOR TASK LIST ticks MON to SUN on that row.
function tsAutoTick_(e) {
  try {
    const r = e && e.range; if (!r) return;
    const sh = r.getSheet(); if (sh.getName() !== tabName_(TS_TAB)) return;
    const hc = ts3Is_(sh) ? TS3.how + 1 : 4;
    if (r.getColumn() > hc || r.getLastColumn() < hc) return;
    for (let row = Math.max(2, r.getRow()); row <= r.getLastRow(); row++) {
      if (/^(every ?day|daily)$/i.test(String(sh.getRange(row, hc).getValue() || '').trim())) sh.getRange(row, hc + 1, 1, 7).setValues([[true, true, true, true, true, true, true]]);
    }
  } catch (x) {}
}
function twSheetEdit(e) {
  try { if (e && e.range && e.range.getSheet().getName() === tabName_(ZEL_FAQ_TAB)) { zelFaqTidy_(); return; } } catch (x) {}
  try { if (e && e.range && e.range.getSheet().getName() === tabName_(TN_TAB)) { tnPlanStatus_(); return; } } catch (x) {}
  try { if (e && e.range && e.range.getSheet().getName() === tabName_(RC_TAB)) { twLater_('rc'); return; } } catch (x) {}
  tsAutoTick_(e);
  try { const n = e && e.range ? e.range.getSheet().getName() : ''; if ([TS_TAB, TC_TAB, PRL_TAB, WV_TAB, INV_TAB, AL_TAB, LBL_TAB].map(tabName_).indexOf(n) < 0) return; } catch (x) { return; }
  twLater_('sync');
}
// Undo-friendly: after an edit the script waits (2 minutes after the LAST edit) before it writes its own Status marks and updates the
// tablets, so Undo in the Sheet only undoes your own changes. "Update tablets with Sheet changes" in the Birriaz menu does it now.
// Script property TW_EDIT_DELAY = seconds to wait (0 = right away).
const TW_ANY = ['ANY CASHIER', 'ANY COOK', 'ANY DISHES'];
function twDelayMs_() { try { const v = PropertiesService.getScriptProperties().getProperty('TW_EDIT_DELAY'); if (v !== null && v !== '' && !isNaN(+v)) return Math.max(0, +v) * 1000; } catch (x) {} return 120000; }
function twLater_(what) {
  const ms = twDelayMs_();
  if (!ms) { if (what === 'rc') rcRead_(true); else twSheetSync_(); return; }
  const P = PropertiesService.getScriptProperties(), now = Date.now();
  P.setProperty('TW_EDIT_AT', String(now));
  if (what === 'rc') P.setProperty('TW_LATER_RC', '1'); else P.setProperty('TW_LATER_SYNC', '1');
  if (now - +(P.getProperty('TW_LATER_AT') || 0) < 20 * 60000 && ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'twSheetLater'; })) return;
  try { ScriptApp.newTrigger('twSheetLater').timeBased().after(ms).create(); P.setProperty('TW_LATER_AT', String(now)); }
  catch (x) { if (what === 'rc') rcRead_(true); else twSheetSync_(); }
}
function twSheetLater() {
  const P = PropertiesService.getScriptProperties(), ms = twDelayMs_();
  try { ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'twSheetLater') ScriptApp.deleteTrigger(t); }); } catch (x) {}
  P.deleteProperty('TW_LATER_AT');
  const wait = ms - (Date.now() - +(P.getProperty('TW_EDIT_AT') || 0));
  if (wait > 5000) { try { ScriptApp.newTrigger('twSheetLater').timeBased().after(wait).create(); P.setProperty('TW_LATER_AT', String(Date.now())); return; } catch (x) {} }
  const rc = P.getProperty('TW_LATER_RC') === '1', sy = P.getProperty('TW_LATER_SYNC') === '1';
  P.deleteProperty('TW_LATER_RC'); P.deleteProperty('TW_LATER_SYNC');
  if (rc) { try { rcRead_(true); } catch (x) {} }
  if (sy || !rc) twSheetSync_();
}
// Writes a Status column only when something in it changed (every script write is one more step for Undo).
function stWrite_(sh, r0, col, vals) {
  if (!vals.length) return false;
  try { const cur = sh.getRange(r0, col, vals.length, 1).getValues(); if (cur.every(function (r, i) { return String(r[0] === null ? '' : r[0]) === String(vals[i][0]); })) return false; } catch (x) {}
  sh.getRange(r0, col, vals.length, 1).setValues(vals); return true;
}
function twSheetSyncMenu() {
  const r = twSheetSync_();
  try { SpreadsheetApp.getActiveSpreadsheet().toast(r.ok ? 'Tablets updated: ' + r.changes + ' change(s). Problems: ' + r.problems : 'Not updated: ' + r.error, 'Digital Pilot OPS', 8); } catch (x) {}
  return r;
}
// Run ONCE from the editor: creates the Tasks and Collaborators tabs filled with today's tasks, and the edit trigger.
// It never overwrites a tab that already has rows.
function setupTaskSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), cfg = twConfig_(), today = logNow_().date, out = [];
  let ct = tabFind_(ss, TC_TAB);
  if (!ct) ct = ss.insertSheet(tabName_(TC_TAB));
  if (ct.getLastRow() < 2) {
    const rows = cfg.employees.map(function (e) { return [e.name, e.role || TS_ROLE[e.name] || '', e.shift || TS_SHIFT[e.name] || '', e.off ? 'No' : 'Yes', '', e.rend || '']; });
    ct.getRange(1, 1, 1, TC_HEAD.length).setValues([TC_HEAD]).setFontWeight('bold');
    ct.getRange(2, 1, rows.length, TC_HEAD.length).setValues(rows);
    try { ct.setFrozenRows(1); ct.setColumnWidth(1, 160); ct.setColumnWidth(5, 260); } catch (x) {}
    out.push('Collaborators tab created (' + rows.length + ' names).');
  } else out.push('Collaborators tab already has names, left as is.');
  let sh = tabFind_(ss, TS_TAB);
  if (!sh) sh = ss.insertSheet(tabName_(TS_TAB));
  if (sh.getLastRow() < 2) {
    const ord = cfg.employees.map(function (e) { return e.name; });
    const list = cfg.tasks.filter(function (t) { return !t.off && !(t.f === 'ONE' && t.date < today); })
      .sort(function (a, b) { return (ord.indexOf(a.rot ? a.rot[0] : a.emp) - ord.indexOf(b.rot ? b.rot[0] : b.emp)) || TS_HOW.indexOf(tsHowTxt_(a)) - TS_HOW.indexOf(tsHowTxt_(b)) || a.id.localeCompare(b.id); });
    const rows = list.map(tsRowOf_);
    sh.getRange(1, 1, 1, TS_HEAD.length).setValues([TS_HEAD]).setFontWeight('bold');
    if (rows.length) sh.getRange(2, 1, rows.length, TS_HEAD.length).setValues(rows);
    try {
      tsFormat_(sh, rows.length);
      sh.getRange(1, 1, Math.max(1, sh.getMaxRows()), 1).setBackground('#EEEEEE');
    } catch (x) {}
    out.push('Tasks tab created (' + rows.length + ' tasks).');
  } else out.push('Tasks tab already has rows, left as is.');
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'twSheetEdit') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('twSheetEdit').forSpreadsheet(ss).onEdit().create();
    out.push('Edit trigger installed.');
  } catch (x) { out.push('Edit trigger NOT installed: ' + x); }
  try { pcfgSetupTabs_(ss, out); } catch (x) { out.push('PREP LIST / WEEK VIEW tabs NOT created: ' + x); }
  try { invSetupTab_(ss, out); } catch (x) { out.push('INVENTORY tab NOT created: ' + x); }
  try { invMergeTabs_(out); } catch (x) { out.push('Old inventory tabs NOT merged: ' + x); }
  invDailyTrigger_(out);
  invAutoTrigger_(out);
  try { alSetupTab_(ss, out); } catch (x) { out.push('ALARMS tab NOT created: ' + x); }
  try { zelFaqSetup_(ss, out); } catch (x) { out.push('ZELYNA Q&A tab NOT created: ' + x); }
  try { tnTriggers_(out); } catch (x) { out.push('Daily training triggers NOT installed: ' + x); }
  try { mrSetupTab_(ss, out); } catch (x) { out.push('ISSUES tab NOT created: ' + x); }
  const r = twSheetSync_(); out.push('First update: ' + JSON.stringify(r));
  try { out.push('Formatted tabs: ' + formatSheets().join(', ')); } catch (x) { out.push('Formatting NOT done: ' + x); }
  Logger.log(out.join('\n'));
  return out;
}

/* ===================== ZELYNA (formerly NOVI): the assistant identity for every automated email ===================== */
// Every observation and recommendation below is calculated from the real records. Nothing is guessed.
// Optional: to let an AI model write Zelyna's notes in its own words, add a Script property named ZELYNA_AI_KEY (or the old NOVI_AI_KEY)
// (Project Settings > Script properties) with an Anthropic API key. Without it Zelyna uses the built-in analysis.
const NOVI_LOC = 'Birriaz – San Jacinto', NOVI_FROM = 'Zelyna - Digital Pilot OPS', NOVI_MODEL = 'claude-haiku-5-5';
function noviRange_(a, b) { const A = logFmt_(a), B = logFmt_(b); return A.slice(0, A.lastIndexOf(',')) + ' – ' + B; }
function noviTime_() { const m = logNow_().min, h = Math.floor(m / 60); return ((h % 12) || 12) + ':' + ('0' + (m % 60)).slice(-2) + (h < 12 ? ' AM' : ' PM'); }
function noviPl_(n, w, p) { return n + ' ' + (n === 1 ? w : (p || w + 's')); }
function noviList_(a, max) { a = a.slice(0, max || 5); return a.length <= 1 ? a.join('') : a.slice(0, -1).join(', ') + ' and ' + a[a.length - 1]; }
function noviAi_(o) {
  let key = '';
  key = zelKey_();
  if (!key) return { used: false };
  try {
    const sys = 'You are Zelyna, the friendly restaurant operations assistant for ' + NOVI_LOC + '. Rewrite the draft observations and recommendations for the manager. ' +
      'Use ONLY the facts given. Never invent numbers, names, records, inventory amounts or problems. If the facts are not enough for a conclusion, say so. ' +
      'Keep it short, specific, friendly and professional. Reply with JSON only: {"observations":["..."],"recommendations":[{"text":"...","why":"..."}]}';
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: NOVI_MODEL, max_tokens: 1200, system: sys, messages: [{ role: 'user', content: JSON.stringify({ report: o.report, period: o.period, summary: o.summary, facts: o.facts || {}, draft_observations: o.obs, draft_recommendations: o.recs }) }] })
    });
    if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
    const j = JSON.parse(res.getContentText()), txt = (j.content || []).map(function (c) { return c.text || ''; }).join(''), m = /\{[\s\S]*\}/.exec(txt);
    const out = JSON.parse(m[0]);
    const obs = (out.observations || []).map(String).filter(Boolean).slice(0, 8);
    const recs = (out.recommendations || []).map(function (r) { return typeof r === 'string' ? { text: r, why: '' } : { text: String(r.text || ''), why: String(r.why || '') }; }).filter(function (r) { return r.text; }).slice(0, 6);
    if (!obs.length) throw new Error('empty answer');
    return { used: true, obs: obs, recs: recs };
  } catch (e) { return { used: false, failed: true, error: String(e) }; }
}
// o = {to, report, period, date, summary:[], obs:[], recs:[{text, why}], detail, attachments, test, facts}
function noviBuild_(o) {
  const ai = noviAi_(o), obs = ai.used ? ai.obs : o.obs, recs = ai.used ? ai.recs : o.recs;
  const subject = (o.test ? 'TEST - ' : '') + 'Zelyna | ' + o.report + ' | ' + NOVI_LOC + ' | ' + o.date;
  const L = ['Hi, this is Zelyna, your restaurant operations assistant from ' + NOVI_LOC + '.', '', "I'm here to help keep your restaurant organized and running smoothly.", '',
    'Here is your ' + o.report + ' for ' + o.period + '.', ''];
  if (o.test) L.push('(This is a TEST email. Nothing was recorded as sent.)', '');
  L.push('REPORT SUMMARY'); (o.summary || []).forEach(function (s) { L.push('- ' + s); }); L.push('');
  L.push("ZELYNA'S OBSERVATIONS");
  if (ai.failed) L.push("(Zelyna's AI analysis is temporarily unavailable, so these notes were calculated directly from the report data.)");
  (obs.length ? obs : ['There is not enough information in this report for me to draw any conclusions.']).forEach(function (s) { L.push('- ' + s); });
  L.push('', "ZELYNA'S RECOMMENDATIONS");
  if (recs.length) recs.forEach(function (r, i) { L.push((i + 1) + '. ' + r.text); if (r.why) L.push('   Why it matters: ' + r.why); });
  else L.push('- Nothing in this report needs action right now.');
  if (o.detail) L.push('', 'REPORT DETAILS', o.detail);
  L.push('', "That's everything for this report!", '', "I'm here to help you stay organized, save time, and keep operations running smoothly.", '',
    '— Zelyna', 'Your Restaurant Operations Assistant', 'Powered by Zelyna · Digital Pilot OPS', NOVI_LOC);
  return { subject: subject, body: L.join('\n'), ai: ai };
}
function noviMail_(o) {
  const m = noviBuild_(o);
  MailApp.sendEmail({ to: o.to, subject: m.subject, body: m.body, attachments: o.attachments || [], name: NOVI_FROM });
  return m;
}
function noviRows_(tab, cols) { try { const sh = tabFind_(SpreadsheetApp.getActiveSpreadsheet(), tab); if (!sh || sh.getLastRow() < 2) return []; return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues(); } catch (e) { return []; } }
function noviDay_(d) { return Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d) ? Utilities.formatDate(d, LOG_TZ, 'yyyy-MM-dd') : ''; }

/* ---- weekly food temperature log ---- */
function noviTemp_(week, st, entries, prev) {
  const sun = logAddDays_(week, 6), obs = [], recs = [], facts = { compliance: st.compliance, done: st.done, total: st.total, missing: st.missing, failed: st.failed, actions: st.actions, discarded: st.discarded };
  const summary = ['Completed checks: ' + st.done + ' of ' + st.total, 'Missed checks: ' + st.missing, 'Checks logged empty (no food): ' + st.empty,
    'Readings below 135°F: ' + st.failed, 'Corrective actions recorded: ' + st.actions, 'Weekly compliance: ' + logCompTxt_(st)];
  if (!st.total) { obs.push("There were no scheduled temperature checks this week, so I can't draw any conclusions yet."); return { summary: summary, obs: obs, recs: recs, facts: facts }; }
  if (prev && prev.total && prev.compliance !== null && st.compliance !== null) {
    const d = Math.round((st.compliance - prev.compliance) * 10) / 10;
    obs.push('Weekly compliance was ' + st.compliance + '%, compared with ' + prev.compliance + '% the week before' + (d ? ' (' + (d > 0 ? 'up ' : 'down ') + Math.abs(d) + ' points).' : ' (no change).'));
    facts.prevCompliance = prev.compliance;
  } else obs.push("There is no earlier week to compare with yet.");
  const missBy = {}, openFail = [];
  st.days.forEach(function (day) { day.slots.forEach(function (s) { if (s.status === 'missing') { const k = LOG_CHECKS[s.ci]; missBy[k] = (missBy[k] || 0) + 1; } if (s.status === 'fail') openFail.push(day.date + ' ' + LOG_CHECKS[s.ci] + ' ' + (s.food || '')); }); });
  const worst = Object.keys(missBy).sort(function (a, b) { return missBy[b] - missBy[a]; })[0];
  if (st.missing) obs.push(noviPl_(st.missing, 'check was', 'checks were') + ' missed.' + (worst ? ' The ' + worst + ' check was missed the most (' + noviPl_(missBy[worst], 'time') + ').' : ''));
  const fails = entries.filter(function (e) { return e.result === 'FAIL'; }), byFood = {};
  fails.forEach(function (e) { const f = e.name || e.food; byFood[f] = (byFood[f] || 0) + 1; });
  if (fails.length) obs.push(noviPl_(fails.length, 'reading was', 'readings were') + ' below 135°F (' + noviList_(Object.keys(byFood).map(function (f) { return f + ' ' + byFood[f] + 'x'; }), 6) + ').');
  if (openFail.length) obs.push(noviPl_(openFail.length, 'low reading has', 'low readings have') + ' no passing follow-up recorded.');
  if (st.discarded) obs.push('Food was discarded ' + noviPl_(st.discarded, 'time') + ' after staying below 135°F.');
  const ed = entries.filter(function (e) { return e.edits && e.edits.length; }).length;
  if (ed) obs.push(noviPl_(ed, 'reading was', 'readings were') + ' edited after saving. The original values are kept in the record.');
  if (openFail.length) recs.push({ text: 'Review the ' + noviPl_(openFail.length, 'low reading') + ' with no passing follow-up and confirm what was done with that food.', why: 'Food held under 135°F for more than 2 hours must be discarded, so every low reading needs a follow-up.' });
  Object.keys(byFood).filter(function (f) { return byFood[f] >= 2; }).slice(0, 2).forEach(function (f) { recs.push({ text: 'Check how ' + f + ' is being held (heat dial, water level, lid). It was below 135°F ' + byFood[f] + ' times this week.', why: 'The same food failing more than once points to a holding or equipment problem, not a one-off.' }); });
  if (worst && missBy[worst] >= 2) recs.push({ text: 'Make sure someone owns the ' + worst + ' check. It was missed ' + missBy[worst] + ' times.', why: 'Missed checks lower compliance and leave no record that the food was safe.' });
  if (!recs.length && st.compliance !== null && st.compliance < 90) recs.push({ text: 'Go over the temperature log routine with the team. Compliance was ' + st.compliance + '%.', why: 'A weekly compliance under 90% means several checks were not recorded.' });
  if (!recs.length && !st.missing && !fails.length) recs.push({ text: 'Keep the current routine. Every scheduled check was completed and no reading was below 135°F.', why: 'Consistent records are what a health inspector looks for.' });
  facts.missedByCheck = missBy; facts.failedByFood = byFood; facts.openFails = openFail.length;
  return { summary: summary, obs: obs, recs: recs, facts: facts };
}

/* ---- weekly collaborator task report ---- */
function noviTasks_(cfg, r) {
  const st = r.st, S = st.store, obs = [], recs = [], tk = {}, facts = { pct: S.pct, due: S.due, done: S.done, overdue: S.overdue };
  cfg.tasks.forEach(function (t) { tk[t.id] = t; });
  const summary = ['Overall completion: ' + twPct_(S.pct) + ' (' + S.done + ' of ' + S.due + ' tasks due)', 'On time: ' + S.onTime + ' | Late: ' + S.lateDone + ' | Incomplete: ' + S.incomplete,
    'Overdue (still open): ' + S.overdue + ' | Urgent carryover: ' + S.rolled];
  st.order.forEach(function (n) { const a = st.byEmp[n]; if (a.due || a.sched) summary.push(n + ': ' + twPct_(a.pct) + ' (' + a.done + ' of ' + a.due + ')'); });
  if (!S.due) { obs.push("No tasks were due this week, so there is nothing to measure yet."); return { summary: summary, obs: obs, recs: recs, facts: facts }; }
  const ws = (r.hist && r.hist.summary) || [], prev = ws.length > 1 ? ws[ws.length - 2].store : null;
  if (prev && prev.due && prev.pct !== null && S.pct !== null) { const d = Math.round((S.pct - prev.pct) * 10) / 10; obs.push('Store completion was ' + S.pct + '%, compared with ' + prev.pct + '% the week before' + (d ? ' (' + (d > 0 ? 'up ' : 'down ') + Math.abs(d) + ' points).' : ' (no change).')); facts.prevPct = prev.pct; }
  else obs.push('There is no earlier week to compare with yet.');
  const ppl = st.order.filter(function (n) { return st.byEmp[n].due > 0; }).sort(function (a, b) { return st.byEmp[b].pct - st.byEmp[a].pct; });
  if (ppl.length >= 2) { const hi = ppl[0], lo = ppl[ppl.length - 1], A = st.byEmp[hi], B = st.byEmp[lo]; if (A.pct !== B.pct) obs.push('Highest completion: ' + hi + ' ' + A.pct + '% (' + A.done + ' of ' + A.due + '). Lowest: ' + lo + ' ' + B.pct + '% (' + B.done + ' of ' + B.due + ').'); }
  const open = st.items.filter(function (x) { return x.st === 'overdue'; }).sort(function (a, b) { return b.late - a.late; });
  if (open.length) { const o = open[0]; obs.push(noviPl_(open.length, 'task is', 'tasks are') + ' still open and overdue. The oldest is "' + ((tk[o.tid] || {}).en || '') + '" (' + o.emp + ', ' + noviPl_(o.late, 'day') + ' overdue).'); }
  const rec = (r.hist && r.hist.recurring) || [];
  if (rec.length) obs.push('Tasks that keep slipping (missed or late 2+ times in the last ' + noviPl_(ws.length, 'week') + '): ' + noviList_(rec.slice(0, 3).map(function (x) { return '"' + ((tk[x.tid] || {}).en || '') + '" (' + (x.emp || '') + ', ' + x.bad + 'x)'; })) + '.');
  const none = cfg.employees.filter(function (e) { return !e.off && !(st.byEmp[e.name] && st.byEmp[e.name].due); }).map(function (e) { return e.name; });
  if (none.length) obs.push(noviList_(none, 6) + (none.length === 1 ? ' had' : ' had') + ' no tasks due this week.');
  ppl.slice().reverse().filter(function (n) { const a = st.byEmp[n]; return a.pct !== null && a.pct < 70 && a.due >= 3; }).slice(0, 3).forEach(function (n) { const a = st.byEmp[n]; recs.push({ text: 'Follow up with ' + n + ': ' + a.done + ' of ' + a.due + ' tasks done (' + a.pct + '%).', why: 'Below 70% means several tasks were not done, and they roll over as URGENT.' }); });
  if (open.length) recs.push({ text: 'Go over the ' + noviPl_(open.length, 'open overdue task') + ' with the team, starting with "' + ((tk[open[0].tid] || {}).en || '') + '".', why: 'Open tasks keep rolling over to the next day and pile up.' });
  if (rec.length) recs.push({ text: 'Consider moving "' + ((tk[rec[0].tid] || {}).en || '') + '" to a different day or person on the Tasks tab of the Google Sheet. It was missed or late ' + rec[0].bad + ' times.', why: 'A task that keeps slipping usually does not fit that shift or day.' });
  if (none.length) recs.push({ text: 'If ' + noviList_(none, 6) + (none.length === 1 ? ' is' : ' are') + ' working, add tasks for them on the Tasks tab.', why: 'Shared work goes faster when everyone on shift has tasks.' });
  if (!recs.length) recs.push({ text: 'Keep it up. Completion was ' + S.pct + '% and nothing is overdue.', why: 'Recognizing a good week helps keep the routine going.' });
  facts.recurring = rec.length; facts.open = open.length; facts.noTasks = none;
  return { summary: summary, obs: obs, recs: recs, facts: facts };
}

/* ---- inventory order ---- */
function noviOrder_(loc, items, past) {
  const obs = [], recs = [], now = new Date(), today = noviDay_(now), facts = { items: items.length };
  const units = items.reduce(function (s, i) { return s + (Number(i.count) || 0); }, 0);
  const summary = ['Items out of stock: ' + items.length, 'Total count: ' + units];
  const groups = {}, keys = [];
  past.forEach(function (r) { if (String(r[1]) !== String(loc)) return; const k = Object.prototype.toString.call(r[0]) === '[object Date]' ? Utilities.formatDate(r[0], LOG_TZ, 'yyyy-MM-dd HH:mm') : String(r[0]); if (!groups[k]) { groups[k] = []; keys.push(k); } groups[k].push({ n: String(r[3]), k: String(r[2] || r[3]).trim().toUpperCase(), c: Number(r[4]) || 0 }); });
  const ik = function (i) { return String(i.code || i.name).trim().toUpperCase(); };
  keys.sort().reverse();
  if (!keys.length) { obs.push("This is the first order on record, so I can't compare it with earlier orders yet."); return { summary: summary, obs: obs, recs: recs, facts: facts }; }
  const lastDay = keys[0].slice(0, 10), ago = twDayNum_(today) - twDayNum_(lastDay);
  obs.push('The last order was sent on ' + logFmt_(lastDay) + ' (' + (ago === 0 ? 'earlier today' : noviPl_(ago, 'day') + ' ago') + ') with ' + noviPl_(groups[keys[0]].length, 'item') + '.');
  const k3 = keys.slice(0, 3), every = items.filter(function (i) { return k3.length >= 3 && k3.every(function (k) { return groups[k].some(function (g) { return g.k === ik(i) || g.n === i.name; }); }); }).map(function (i) { return i.name; });
  if (every.length) obs.push('Also out of stock in each of the last ' + k3.length + ' orders: ' + noviList_(every, 6) + '.');
  const lim = logAddDays_(today, -60), recent = keys.filter(function (k) { return k.slice(0, 10) >= lim; });
  const fresh = items.filter(function (i) { return !recent.some(function (k) { return groups[k].some(function (g) { return g.k === ik(i) || g.n === i.name; }); }); }).map(function (i) { return i.name; });
  if (fresh.length && fresh.length < items.length) obs.push('Not on any order in the last 60 days: ' + noviList_(fresh, 6) + '.');
  const big = [];
  items.forEach(function (i) { const h = []; keys.forEach(function (k) { groups[k].forEach(function (g) { if ((g.k === ik(i) || g.n === i.name) && g.c > 0) h.push(g.c); }); }); if (h.length >= 2) { const avg = h.reduce(function (s, x) { return s + x; }, 0) / h.length; if ((Number(i.count) || 0) >= 2 * avg && i.count - avg >= 2) big.push({ n: i.name, c: i.count, a: Math.round(avg * 10) / 10 }); } });
  if (big.length) obs.push('Higher than usual: ' + noviList_(big.map(function (b) { return b.n + ' ' + b.c + ' (usually about ' + b.a + ')'; }), 4) + '.');
  if (every.length) recs.push({ text: 'Consider ordering a larger amount of ' + noviList_(every, 4) + ' so it lasts until the next delivery.', why: 'It ran out before each of the last ' + k3.length + ' orders, so the current amount is not enough.' });
  big.slice(0, 2).forEach(function (b) { recs.push({ text: 'Double-check the count for ' + b.n + ' (' + b.c + ', usually about ' + b.a + ') before the order is placed.', why: 'A count far above normal is often a typing mistake.' }); });
  facts.repeat = every; facts.notRecent = fresh; facts.higher = big; facts.daysSinceLast = ago;
  return { summary: summary, obs: obs, recs: recs, facts: facts };
}

/* ---- maintenance / damage report ---- */
const NOVI_EQUIP = ['fryer', 'freidora', 'fridge', 'refrigerador', 'refri', 'freezer', 'congelador', 'cooler', 'hielera', 'steam table', 'mesa de vapor', 'warmer', 'calentador', 'a/c', 'aire', 'sink', 'tarja', 'lavamanos', 'drain', 'drenaje', 'toilet', 'baño', 'door', 'puerta', 'light', 'luz', 'printer', 'impresora', 'grill', 'plancha', 'griddle', 'hood', 'campana', 'ice', 'hielo', 'soda', 'pepsi', 'register', 'caja', 'water heater', 'boiler', 'faucet', 'llave', 'oven', 'horno', 'stove', 'estufa', 'roof', 'techo', 'floor', 'piso', 'leak', 'fuga'];
function noviMaint_(loc, issue, photos, past) {
  const obs = [], recs = [], low = String(issue).toLowerCase(), today = logNow_().date;
  const summary = ['Location: ' + loc, 'Photos attached: ' + photos, 'Issue: ' + String(issue).replace(/\s+/g, ' ').slice(0, 200) + (String(issue).length > 200 ? '...' : '')];
  const safety = /\b(gas|smoke|humo|fire|fuego|spark|chispa|burning|quemad|shock|descarga|electric|el[eé]ctric)/.test(low);
  const food = /(fridge|refrigerador|refri|freezer|congelador|cooler|hielera|steam table|mesa de vapor|warmer|calentador)/.test(low);
  const water = /(leak|fuga|flood|inund|water|agua)/.test(low);
  const kw = NOVI_EQUIP.filter(function (w) { return low.indexOf(w) >= 0; });
  if (safety) obs.push('This report mentions a possible safety hazard (' + (low.match(/\b(gas|smoke|humo|fire|fuego|spark|chispa|burning|shock|electric\w*|el[eé]ctric\w*)/) || [''])[0] + ').');
  if (food) obs.push('The equipment involved holds food at a safe temperature, so food could be at risk.');
  const last90 = logAddDays_(today, -90), sim = past.filter(function (r) { const d = noviDay_(r[0]); const t = String(r[2] || '').toLowerCase(); return d >= last90 && kw.some(function (w) { return t.indexOf(w) >= 0; }); });
  if (kw.length && sim.length) obs.push('This is report #' + (sim.length + 1) + ' in the last 90 days that mentions ' + noviList_(kw.slice(0, 3).map(function (w) { return '"' + w + '"'; })) + ' (last one on ' + logFmt_(noviDay_(sim[sim.length - 1][0])) + ').');
  else if (kw.length) obs.push('No other report in the last 90 days mentions ' + noviList_(kw.slice(0, 3).map(function (w) { return '"' + w + '"'; })) + '.');
  const n30 = past.filter(function (r) { return noviDay_(r[0]) >= logAddDays_(today, -30); }).length;
  obs.push((n30 ? noviPl_(n30, 'other maintenance report') : 'No other maintenance reports') + ' in the last 30 days.');
  if (!photos) obs.push('No photo was attached.');
  if (safety) recs.push({ text: 'Treat this as urgent: stop using the equipment if it is unsafe and call a technician today.', why: 'Gas, smoke, sparks or electrical problems can injure people or start a fire.' });
  if (food) recs.push({ text: 'Check and log the temperature of the food in this unit now, and move the food if it is not holding temperature.', why: 'Food that stays out of safe temperature has to be thrown away.' });
  if (water && !safety) recs.push({ text: 'Put a wet floor sign out and keep the area dry until it is fixed.', why: 'Water on the floor is a slip hazard for the team and customers.' });
  if (kw.length && sim.length) recs.push({ text: 'Since this has come up before, ask the technician for a lasting repair or a replacement quote.', why: 'Repeat problems cost more over time than one proper fix.' });
  if (!photos) recs.push({ text: 'Next time, attach a photo so the technician can see the problem before coming.', why: 'A photo often saves a trip and speeds up the repair.' });
  return { summary: summary, obs: obs, recs: recs, facts: { safety: safety, food: food, keywords: kw, similar90: sim.length, reports30: n30, photos: photos } };
}

/* ---- close out checklist ---- */
function noviClose_(loc, items, notes, photos, past) {
  const obs = [], recs = [], now = new Date(), today = noviDay_(now), dow = new Date(today + 'T12:00:00Z').getUTCDay();
  const alarm = (dow >= 1 && dow <= 3) ? '20:13' : '21:48', am = +alarm.slice(0, 2) * 60 + +alarm.slice(3), nm = logNow_().min;
  const summary = ['Items checked: ' + items.length, 'Sent at: ' + noviTime_(), 'Photos attached: ' + photos, 'Notes: ' + (notes ? 'yes' : 'none')];
  const late = nm - am;
  obs.push('Sent at ' + noviTime_() + (late > 30 ? ', ' + late + ' minutes after the close out alarm.' : late >= 0 ? ', right after the close out alarm.' : ', before the close out alarm.'));
  const days = {}; past.forEach(function (r) { const d = noviDay_(r[0]); if (d) days[d] = 1; }); days[today] = 1;
  const miss = []; for (let i = 1; i <= 6; i++) { const d = logAddDays_(today, -i); if (!days[d]) miss.push(d); }
  const anyPast = past.length > 0;
  if (anyPast) obs.push('Close outs were sent on ' + (7 - miss.length) + ' of the last 7 days' + (miss.length ? ' (none on ' + noviList_(miss.map(logFmt_), 6) + ').' : '.'));
  else obs.push("This is the first close out on record, so I can't compare it with earlier days yet.");
  if (notes) obs.push('The closing team left notes.');
  if (notes) recs.push({ text: 'Read the closing notes and follow up on anything that needs action tomorrow.', why: 'Notes are often the only record of a problem found at closing.' });
  if (anyPast && miss.length) recs.push({ text: 'Check why no close out was sent on ' + noviList_(miss.map(logFmt_), 3) + '.', why: 'Every night needs a close out so gas, equipment, water and doors are confirmed off and locked.' });
  if (late > 60) recs.push({ text: 'Remind the closing team to send the close out when the alarm rings.', why: 'A late close out can mean something was left on longer than needed.' });
  return { summary: summary, obs: obs, recs: recs, facts: { late: late, missingDays: miss, notes: !!notes, photos: photos } };
}

/* ===================== PREP LIST + WEEK VIEW from the Google Sheet ===================== */
// The PREP LIST and WEEK VIEW tabs are the place to change the prep items and the weekly prep plan.
// They are read with the Tasks tab (same edit trigger, same 30-minute check) and reach the tablets through the shared prep list sync.
const PRL_TAB = 'PREP LIST', WV_TAB = 'WEEK VIEW', PCFG_FILE = 'prep-config.json';
const PRL_HEAD = ['Item', 'Category', 'Who does it (label)', 'Shelf life (days)', 'Order', 'Active', 'Status (filled in automatically)'];
const WV_HEAD = ['Day', 'Category', 'Item', 'Who does it (label)', 'Order', 'Active', 'Start date', 'Status (filled in automatically)'];
const WV_DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const PLBL_SEED = {'ASADA':'PM COOK','POLLO':'PM COOK','ALPASTOR':'PM COOK','PICAR ASADA':'PM COOK','PICAR POLLO PM':'PM COOK','PICAR AL PASTOR':'PM COOK',
'CILANTRO':'MILTZY','WASH CILANTRO TODO':'MILTZY','PICAR CILANTRO':'MILTZY','PICAR CILANTRO TODO':'MILTZY','LIMONES':'HUGO','LIMONES TODO':'HUGO','CEBOLLA BLANCA':'ERIKA','WHITE ONION TODO':'ERIKA','CEBOLLA MORADA':'ERIKA','PURPLE ONION TODO':'ERIKA',
'TOMATO':'MILTZY','TOMATE 1 CONT':'MILTZY','PICO':'MILTZY','PICO 1 CONT':'MILTZY','FAJITAS':'HUGO','FAJITAS 2-3 CONT':'HUGO','REPOLLO':'AM COOK','LECHUGA':'HUGO','LECHUGA 2 CONT':'HUGO','CEBOLLA PARA ALPASTOR':'HUGO','CEBOLLA ALPASTOR 1 CONT':'HUGO',
'ARROS BLANCO':'AM COOK','WHITE RICE':'AM COOK','ARROS MEXICANO':'ERIKA','MEXICAN RICE':'ERIKA','NACHOS':'ERIKA','FRIJOLES':'HUGO','BEANS':'HUGO','ACEITE ROJO':'HUGO','GRILL CHICKEN':'HUGO','GRILL ALPASTOR':'HUGO',
'SALSA ROJA':'ERIKA','SALSA VERDE':'ERIKA','SALSAS':'ERIKA','ROSTISAR CHILES':'CASHIER','ROAST CHILES':'CASHIER','POTATO':'ERIKA','TAMALES':'ERIKA','TAMALES PREP':'ERIKA','TAMALES COOK':'ERIKA',
'BATIDAS DE ARROZ':'1 DAY BEFORE MEXICAN RICE','MARINAR ASADA':'ERIKA','MARINAR AL PASTOR':'ERIKA','MARINAR ALPASTOR':'ERIKA','GUARDAR ACEITE':'PM SHIFT','GUARDAR FRIJOLES':'PM SHIFT','PAPEL EN CUADRITOS':'MILTZY'};
function pcfgLoad_() { const c = logReadJson_(PCFG_FILE); if (c && c.items && c.items.length) { pApplyCfg_(c); return c; } return null; }
function pcfgHash_(o) { const s = JSON.stringify(o); let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; } return h.toString(36) + s.length.toString(36); }
function pcfgActive_(v) { return !/^(no|n|false|0|off|inactive|inactivo)$/i.test(String(v === false ? 'no' : v === null || v === undefined ? '' : v).trim()); }
function wvDays_(s) {
  const p = tsPlain_(s).trim(), out = [];
  if (/^(EVERY ?DAY|DAILY|DIARIO|TODOS)/.test(p)) return [0, 1, 2, 3, 4, 5, 6];
  p.split(/[\s,\/;&]+/).forEach(function (w) {
    const k = w.slice(0, 3); let i = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'].indexOf(k); if (i < 0) i = ['DOM', 'LUN', 'MAR', 'MIE', 'JUE', 'VIE', 'SAB'].indexOf(k);
    if (i >= 0 && out.indexOf(i) < 0) out.push(i);
  });
  return out;
}
// Reads both tabs. Returns {cfg, plStatus, wvStatus, err}. Never returns an empty prep list (that would wipe the tablets).
function pcfgFromSheet_(ss) {
  const pl = tabFind_(ss, PRL_TAB), wv = tabFind_(ss, WV_TAB);
  if (!pl || !wv) return { err: 'no tabs' };
  const today = logNow_().date, items = [], seen = {}, lbl = {}, sl = {}, plSt = [], wvSt = [];
  if (pl3Is_(pl)) { const r3 = pl3Parse_(pl, today); return { cfg: r3.cfg, err: r3.err, plStatus: r3.st, pl2: true, pl3: true }; }
  if (pl2Is_(pl)) { const r2 = pl2Parse_(pl, today); return { cfg: r2.cfg, err: r2.err, plStatus: r2.st, pl2: true }; }
  const pr = pl.getLastRow() > 1 ? pl.getRange(2, 1, pl.getLastRow() - 1, 6).getValues() : [];
  const rows = pr.map(function (r, i) { return { r: r, i: i, o: r[4] === '' || r[4] === null || isNaN(+r[4]) ? 1e9 + i : +r[4] }; });
  rows.forEach(function (x) { plSt[x.i] = ['']; });
  rows.slice().sort(function (a, b) { return a.o - b.o || a.i - b.i; }).forEach(function (x) {
    const r = x.r, n = String(r[0] || '').trim().toUpperCase().slice(0, 60), g = String(r[1] || '').trim().toUpperCase().slice(0, 60);
    if (!n && !g && r.every(function (v) { return String(v === null ? '' : v).trim() === ''; })) return;
    const er = [];
    if (!n) er.push('Item is empty'); if (!g) er.push('Category is empty');
    if (n && seen[n]) er.push('This item is listed twice');
    let days = 0; const sv = String(r[3] === null ? '' : r[3]).trim();
    if (sv !== '') { days = Math.round(+sv); if (!(days >= 1 && days <= 60)) er.push('Shelf life must be a number of days from 1 to 60 (leave it empty for the standard 6-day rule)'); }
    if (er.length) { plSt[x.i] = ['⚠ ' + er.join('; ') + ' (not used)']; return; }
    seen[n] = 1;
    if (!pcfgActive_(r[5])) { plSt[x.i] = ['Inactive (hidden on the tablets)']; return; }
    items.push({ n: n, g: g, lbl: String(r[2] || '').trim().toUpperCase().slice(0, 30), sl: days || 0 });
    if (items[items.length - 1].lbl) lbl[n] = items[items.length - 1].lbl; if (days) sl[n] = days;
    plSt[x.i] = ['✓ OK'];
  });
  let weekly = [[], [], [], [], [], [], []], every = [];
  const grid = wvIsGrid_(wv);
  if (grid) {
    const gp = wvParseGrid_(wv, today); weekly = gp.weekly; every = gp.every; gp.st.forEach(function (s, i) { wvSt[i] = s; });
    Object.keys(gp.lbl).forEach(function (n) { if (!lbl[n]) lbl[n] = gp.lbl[n]; });
  } else {
    const wr = wv.getLastRow() > 1 ? wv.getRange(2, 1, wv.getLastRow() - 1, 7).getValues() : [];
    const wrows = wr.map(function (r, i) { return { r: r, i: i, o: r[4] === '' || r[4] === null || isNaN(+r[4]) ? 1e9 + i : +r[4] }; });
    wrows.forEach(function (x) { wvSt[x.i] = ['']; });
    wrows.slice().sort(function (a, b) { return a.o - b.o || a.i - b.i; }).forEach(function (x) {
      const r = x.r; if (r.every(function (v) { return String(v === null ? '' : v).trim() === ''; })) return;
      const day = String(r[0] || '').trim(), g = String(r[1] || '').trim().toUpperCase().slice(0, 60), n = String(r[2] || '').trim().toUpperCase().slice(0, 60), lb = String(r[3] || '').trim().toUpperCase().slice(0, 30), er = [];
      if (!g) er.push('Category is empty'); if (!n) er.push('Item is empty');
      const ev = /(?:every|cada)\s*(\d+)\s*(?:days?|d[ií]as?)/i.exec(day);
      let ds = [], an = '';
      if (ev) { const k = +ev[1]; if (!(k >= 2 && k <= 30)) er.push('Every N days must be 2 to 30'); an = tsDate_(r[6]); if (an === null) er.push('Start date not understood (use 11/3/2026)'); if (!an) an = today; }
      else { ds = wvDays_(day); if (!ds.length) er.push('Day not understood. Use Monday ... Sunday, several days like "Mon, Thu", Every day, or Every 3 days'); }
      if (er.length) { wvSt[x.i] = ['⚠ ' + er.join('; ') + ' (not used)']; return; }
      if (!pcfgActive_(r[5])) { wvSt[x.i] = ['Inactive (hidden on the tablets)']; return; }
      if (lb && !lbl[n]) lbl[n] = lb;
      if (ev) every.push({ c: g, n: n, e: +ev[1], an: an }); else ds.forEach(function (d) { weekly[d].push([g, n]); });
      wvSt[x.i] = ['✓ OK'];
    });
  }
  if (!items.length) return { err: 'The PREP LIST tab has no active items, so nothing was changed on the tablets.', plStatus: plSt, wvStatus: wvSt, wvGrid: grid };
  const cfg = { items: items, weekly: weekly, every: every, lbl: lbl, sl: sl };
  cfg.v = pcfgHash_(cfg);
  return { cfg: cfg, plStatus: plSt, wvStatus: wvSt, wvGrid: grid };
}
// Called from twSheetSync_ (edits, menu and every 30 minutes). Writes prep-config.json only when something changed.
function pcfgSync_(ss) {
  const notes = [];
  try { pl2MigrateIfOld_(ss, notes); } catch (x) { notes.push('PREP LIST NOT changed to the new layout: ' + x); }
  try { pl2StageMerge_(ss, notes); } catch (x) { notes.push('PREP LIST stage column NOT changed: ' + x); }
  try { pl2DropShelf_(ss, notes); } catch (x) { notes.push('PREP LIST shelf life column NOT removed: ' + x); }
  try { layout3_(ss, notes); } catch (x) { notes.push('New layout NOT made: ' + x); }
  const r = pcfgFromSheet_(ss);
  if (r.err === 'no tabs') return { ok: false, error: 'no PREP LIST / WEEK VIEW tabs yet (run setupSheets once)' };
  const stamp = twFmtTs_(logStamp_());
  try {
    const pl = tabFind_(ss, PRL_TAB), wv = tabFind_(ss, WV_TAB);
    if (r.pl2) {
      const stc = r.pl3 ? PL3.st + 1 : PL2_HEAD.length;
      if (r.plStatus && r.plStatus.length) pl.getRange(2, stc, r.plStatus.length, 1).setValues(r.plStatus);
      pl.getRange(1, stc).setValue('Status (filled in automatically) - last check ' + stamp + (r.err ? ' - ⚠ ' + r.err : ''));
      if (r.cfg && wv) { try { if (pl2View_(wv, r.cfg)) tabLook_(wv); } catch (x) {} }
    } else {   // old layout (only if the change to the new layout could not be made)
    if (r.plStatus && r.plStatus.length) pl.getRange(2, 7, r.plStatus.length, 1).setValues(r.plStatus);
    if (r.wvGrid) {
      if (r.wvStatus && r.wvStatus.length) wv.getRange(2, 10, r.wvStatus.length, 1).setValues(r.wvStatus.map(function (s) { return s || ['']; }));
      // One time: tabs converted from the old list layout kept its dropdowns (Day, Category, Active = column F) -> "Invalid" marks. Remove them.
      const P = PropertiesService.getScriptProperties();
      if (!P.getProperty('WV_DV_CLEAN')) { try { wv.getRange(1, 1, Math.max(1, wv.getMaxRows()), 10).clearDataValidations(); P.setProperty('WV_DV_CLEAN', '1'); } catch (x) {} }
    }
    else if (!r.err && r.cfg) { wvWriteGrid_(wv, r.cfg.weekly, r.cfg.every, r.cfg.lbl, r.cfg.items.map(function (i) { return i.n; })); tabLook_(wv); }   // old list layout -> weekly grid
    else if (r.wvStatus && r.wvStatus.length) wv.getRange(2, 8, r.wvStatus.length, 1).setValues(r.wvStatus);
    pl.getRange(1, 7).setValue('Status (filled in automatically) - last check ' + stamp + (r.err ? ' - ⚠ ' + r.err : ''));
    wv.getRange(1, wvIsGrid_(wv) ? 10 : 8).setValue('Status (filled in automatically) - last check ' + stamp);
    }
  } catch (x) {}
  if (r.err) return { ok: false, error: r.err, notes: notes };
  try {
    if (!tabFind_(ss, LBL_TAB)) lblSetup_(ss, r.cfg, notes);
    const L = lblRead_(ss, r.cfg);
    if (L) { delete r.cfg.v; r.cfg.labels = L; r.cfg.v = pcfgHash_(r.cfg); }
  } catch (x) { notes.push('FIFO tab not read: ' + x); }
  const old = logReadJson_(PCFG_FILE);
  if (old && old.v === r.cfg.v) return { ok: true, changed: false, v: r.cfg.v, notes: notes };
  r.cfg.at = logStamp_();
  logWriteJson_(PCFG_FILE, r.cfg);
  return { ok: true, changed: true, v: r.cfg.v, items: r.cfg.items.length, notes: notes };
}
// Fills the two tabs with today's built-in lists (only if they are empty).
function pcfgSetupTabs_(ss, out) {
  let pl = tabFind_(ss, PRL_TAB);
  if (!pl) pl = ss.insertSheet(tabName_(PRL_TAB));
  const wv0 = tabFind_(ss, WV_TAB);
  if (pl.getLastRow() < 2 && wv0 && wv0.getLastRow() > 1 && String(wv0.getRange(1, 10).getValue() || '') !== WV_VIEW_NOTE) {
    // an older WEEK VIEW exists: write the built-in list in the old layout, the first update then combines both into the new layout
    const rows = []; let o = 0;
    PREPS.forEach(function (g) { g[1].forEach(function (n) { o += 10; rows.push([n, g[0], PLBL_SEED[n] || '', '', o, 'Yes', '']); }); });
    pl.getRange(1, 1, 1, PRL_HEAD.length).setValues([PRL_HEAD]); pl.getRange(2, 1, rows.length, PRL_HEAD.length).setValues(rows);
    out.push('PREP LIST tab created (it is combined with the WEEK VIEW days on the first update).');
  } else if (pl.getLastRow() < 2) {
    const cat = []; PREPS.forEach(function (g) { g[1].forEach(function (n) { cat.push({ n: n, g: g[0], lbl: PLBL_SEED[n] || '', sl: 0, active: true }); }); });
    let names = []; try { names = (twConfig_().employees || []).filter(function (e) { return !e.off; }).map(function (e) { return e.name; }); } catch (x) {}
    const b = pl2Build_(cat, WEEKLY, PEVERY.map(function (x) { return { c: x.c, n: x.n, e: x.e, an: x.an }; }), PLBL_SEED, {}, names);
    pl2Write_(pl, b.list, names);
    out.push('PREP LIST tab created (' + b.list.length + ' rows, days included).');
  } else out.push('PREP LIST tab already has rows, left as is.');
  let wv = tabFind_(ss, WV_TAB);
  if (!wv) { wv = ss.insertSheet(tabName_(WV_TAB)); out.push('WEEK VIEW tab created (view only, filled in from the PREP LIST).'); }
  getTab_(FIFO_TAB, FIFO_HEAD); out.push('FIFO REPORT tab ready.');
}
/* ===================== TODAY'S TASK - PREP LIST: ONE tab for prep items, stages and days ===================== */
// One row = one prep item (a stage is part of the name: "PICAR ASADA", "MARINAR ASADA").
// Today's task shows: PRODUCT + NOTE (for example "ASAR ASADA AM 11") in the section of column C.
// Column B = the category of the WEEK VIEW tab. Empty = the item is on Today's task but not on the week view.
// FIFO label names come only from the FIFO tab (column A there), never from this tab.
// The WEEK VIEW tab is made from this tab automatically (view only).
const PL2_HEAD = ['Product name', 'WEEK VIEW category (empty = not on the week view)', "Section on Today's task", 'Who does it', 'Note (1 CONT, AM 11...)',
  'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN', 'Every (days)', 'Start date (for Every)', 'FIFO label', 'Active', 'Status (filled in automatically)'];
const PL2_ROLES = ['AM COOK', 'PM COOK', 'CASHIER', 'AM SHIFT', 'PM SHIFT'];
const PL2_SPARE = 20;
function pl2Is_(sh) { return !!sh && (String(sh.getRange(1, 6).getValue() || '').trim().toUpperCase() === 'MON' || pl3Is_(sh)); }   // new layouts (not the very old one)
function pl2Name_(p, s, note) { return [s, p, note].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().slice(0, 60); }
function pl2Label_(i) { return i.p ? i.p + (i.s ? ' – ' + i.s : '') : i.n; }
function pl2Blank_(v) { return String(v === null || v === undefined || v === false || v === true ? '' : v).trim() === ''; }
// Reads the tab. Returns {cfg, st} (st = status per row from row 2).
function pl2Parse_(sh, today) {
  const n = sh.getLastRow() - 1, rows = n > 0 ? sh.getRange(2, 1, n, 16).getValues() : [], st = [], items = [], weekly = [[], [], [], [], [], [], []], every = [], lbl = {}, sl = {}, seen = {};
  rows.forEach(function (r) {
    const days = r.slice(5, 12).map(alOnCell_);
    if ([0, 1, 2, 3, 4, 12, 13].every(function (k) { return pl2Blank_(r[k]); }) && !days.some(Boolean)) { st.push(['']); return; }
    const p = String(r[0] || '').trim().toUpperCase().replace(/\s+/g, ' ').slice(0, 50), s = '',
      g = String(r[2] || '').trim().toUpperCase().slice(0, 60), who = String(r[3] || '').trim().toUpperCase().slice(0, 30), note = String(r[4] || '').trim().toUpperCase().slice(0, 40);
    let wc = String(r[1] || '').trim().toUpperCase().replace(/\s+/g, ' ').slice(0, 60), wcBad = false;
    if (wc && WV_CATORD.indexOf(wc) < 0) { wcBad = true; wc = ''; }
    const nm = pl2Name_(p, s, note), er = [];
    if (!p) er.push('Product is empty'); if (!g) er.push('Section is empty');
    if (p && seen[nm]) er.push('Same product + note is on another row');
    let ev = 0, an = '';
    if (!pl2Blank_(r[12])) { ev = Math.round(+r[12]); if (!(ev >= 2 && ev <= 30)) er.push('Every (days) must be 2 to 30'); an = tsDate_(r[13]); if (an === null) er.push('Start date not understood (use 11/3/2026)'); }
    const days0 = 0;   // shelf life is on the FIFO tab
    if (er.length) { st.push(['⚠ ' + er.join('; ') + ' (not used)']); return; }
    seen[nm] = 1;
    if (!pcfgActive_(r[15])) { st.push(['Inactive (hidden on the tablets)']); return; }
    const fifo = pl2Blank_(r[14]) && r[14] !== false ? true : alOnCell_(r[14]);
    items.push({ n: nm, g: g, lbl: who, sl: days0 || 0, p: p, s: s, f: fifo, wc: wc });
    if (who) lbl[nm] = who; if (days0) sl[nm] = days0;
    days.forEach(function (on, k) { if (on) weekly[WV_COLS[k]].push([g, nm, wc]); });
    if (ev) every.push({ c: g, n: nm, e: ev, an: an || today, wc: wc });
    const when = days.map(function (on, k) { return on ? WV_HDR[k + 1].slice(0, 3) : ''; }).filter(Boolean).join(', ');
    st.push(['✓ ' + nm + (when ? ' · ' + when : '') + (ev ? ' · every ' + ev + ' days' : '') + (!when && !ev ? ' · any day' : '') + (wcBad ? ' · ⚠ week view category not in the list (not on the week view)' : wc ? ' · week view: ' + wc : ' · not on the week view')]);
  });
  const cfg = { items: items, weekly: weekly, every: every, lbl: lbl, sl: sl };
  cfg.v = pcfgHash_(cfg);
  return { cfg: items.length ? cfg : null, st: st, err: items.length ? '' : 'The PREP LIST tab has no active items, so nothing was changed on the tablets.' };
}
// Writes the tab from a row list [{p,s,g,who,note,days:[dow...],ev,an,f,sl,active}]
function pl2Write_(sh, list, names) {
  try { sh.getBandings().forEach(function (b) { b.remove(); }); } catch (x) {}
  try { sh.getRange(1, 1, Math.max(1, sh.getMaxRows()), Math.max(17, sh.getMaxColumns ? sh.getMaxColumns() : 17)).breakApart().clearDataValidations(); } catch (x) {}
  sh.clear();
  const rows = list.map(function (x) {
    return [(x.s ? x.s + ' ' : '') + x.p, x.wc !== undefined ? x.wc : (((x.days || []).length || x.ev) ? x.g : ''), x.g, x.who || '', x.note || ''].concat(WV_COLS.map(function (d) { return (x.days || []).indexOf(d) >= 0; }))
      .concat([x.ev || '', x.ev ? (x.an || '') : '', x.f !== false, x.active === false ? 'No' : 'Yes', '']);
  });
  for (let i = 0; i < PL2_SPARE; i++) rows.push(['', '', '', '', '', false, false, false, false, false, false, false, '', '', true, '', '']);
  sh.getRange(1, 1, 1, PL2_HEAD.length).setValues([PL2_HEAD]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, PL2_HEAD.length).setValues(rows);
  pl2Format_(sh, rows.length, names);
}
function pl2Format_(sh, n, names) {
  const R = function (c, k) { return sh.getRange(2, c, n, k || 1); };
  try { R(6, 7).insertCheckboxes(); R(15).insertCheckboxes(); } catch (x) {}
  try {
    sh.setFrozenRows(1); sh.setFrozenColumns(2);
    [190, 150, 210, 140, 160, 58, 58, 58, 58, 58, 58, 58, 90, 120, 80, 70, 420].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange(1, 1, 1, PL2_HEAD.length).setWrap(true).setVerticalAlignment('middle');
    sh.getRange(1, 6, n + 1, 7).setHorizontalAlignment('center'); sh.getRange(1, 15, n + 1, 1).setHorizontalAlignment('center');
    R(14).setNumberFormat('m/d/yyyy');
    const dv = function (list) { return SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build(); };
    R(3).setDataValidation(dv(WV_CATORD.slice()));
    R(4).setDataValidation(dv((names || []).concat(PL2_ROLES)));
    R(16).setDataValidation(dv(['Yes', 'No']));
    R(2).setDataValidation(dv(WV_CATORD.slice()));
  } catch (x) {}
}

/* ---- one time: the old PREP LIST (catalog) + WEEK VIEW (days) become the new combined tab ---- */
const PL2_MEAT = { 'ASADA': 'ASADA', 'POLLO': 'POLLO', 'CHICKEN': 'POLLO', 'ALPASTOR': 'ALPASTOR', 'AL PASTOR': 'ALPASTOR', 'PASTOR': 'ALPASTOR' };
const PL2_STAGE = { 'PICAR': 'PICAR', 'ASAR': 'ASAR', 'GRILL': 'ASAR', 'COOK': 'ASAR', 'MARINAR': 'MARINAR' };
// "PICAR POLLO PM" -> {p:POLLO, s:PICAR, note:PM}; "COOK CHICKEN AM 11" -> {p:POLLO, s:ASAR, note:AM 11}; "PICO 1 CONT" -> {p:PICO, note:1 CONT}
function pl2Split_(name, section) {
  let s = String(name || '').trim().toUpperCase().replace(/\s+/g, ' '), note = '';
  const m = /\s+((?:AM|PM)(?:\s+\d{1,2}(?::\d{2})?)?|TODO|\d+(?:-\d+)?\s+CONT)$/.exec(s);
  if (m && m.index > 0) { note = m[1]; s = s.slice(0, m.index).trim(); }
  const w = s.split(' ');
  let st = '', p = s;
  if (w.length >= 2 && PL2_STAGE[w[0]] && PL2_MEAT[w.slice(1).join(' ')]) { st = PL2_STAGE[w[0]]; p = PL2_MEAT[w.slice(1).join(' ')]; }
  else if (w.length >= 2 && PL2_STAGE[w[w.length - 1]] && PL2_MEAT[w.slice(0, -1).join(' ')]) { st = PL2_STAGE[w[w.length - 1]]; p = PL2_MEAT[w.slice(0, -1).join(' ')]; }
  else if (PL2_MEAT[s] && /^PICAR/.test(String(section || '').toUpperCase())) { st = 'PICAR'; p = PL2_MEAT[s]; }
  return { p: p, s: st, note: note };
}
// catalog = [{n,g,lbl,sl,active}], weekly = [[ [g,n], ...] Sun..Sat], every = [{c,n,e,an}], lbl = {name: who}
function pl2Build_(catalog, weekly, every, lbl, linkOwner, names) {
  const list = [], byKey = {}, byBase = {}, map = {};
  const known = {}; (names || []).concat(PL2_ROLES).forEach(function (x) { known[x] = 1; });
  const key = function (o) { return o.s + '|' + o.p.replace(/\s+/g, '') + '|' + o.note; }, base = function (o) { return o.s + '|' + o.p.replace(/\s+/g, ''); };
  const add = function (o, g, who) {
    const r = { p: o.p, s: o.s, note: o.note, g: g, who: '', days: [], ev: 0, an: '', f: true, sl: 0, active: true, old: [] };
    list.push(r); byKey[key(o)] = r; if (!byBase[base(o)]) byBase[base(o)] = r; setWho(r, who); return r;
  };
  const setWho = function (r, who) {
    who = String(who || '').trim().toUpperCase(); if (!who || r.who) return;
    if (!known[who] && who.length > 12) { if (!r.note) r.note = who.slice(0, 40); return; }   // a long "label" like "1 DAY BEFORE MEXICAN RICE" is a note, not a person
    r.who = who;
  };
  const find = function (o) {
    let r = byKey[key(o)]; if (r) return r;
    const b = byBase[base(o)];
    if (o.note && b && !b.note && !b.days.length && !b.ev) { b.note = o.note; byKey[key(o)] = b; return b; }
    if (!o.note && b) return b;
    return null;
  };
  catalog.forEach(function (c) {
    const o = pl2Split_(c.n, c.g); let r = byKey[key(o)];
    if (!r) { r = add(o, c.g, c.lbl); r.sl = c.sl || 0; r.active = c.active !== false; }
    r.old.push(c.n);
  });
  WV_COLS.forEach(function (d) {
    (weekly[d] || []).forEach(function (w) {
      const o = pl2Split_(w[1], w[0]); let r = find(o);
      if (!r) r = add(o, w[0], ''); setWho(r, lbl[w[1]]);
      if (r.days.indexOf(d) < 0) r.days.push(d); if (r.old.indexOf(w[1]) < 0) r.old.push(w[1]);
    });
  });
  (every || []).forEach(function (x) {
    const o = pl2Split_(x.n, x.c); let r = find(o);
    if (!r) r = add(o, x.c, ''); setWho(r, lbl[x.n]); r.ev = x.e; r.an = x.an || ''; if (r.old.indexOf(x.n) < 0) r.old.push(x.n);
  });
  list.forEach(function (r) {
    if (!r.who) r.old.some(function (n) { const w = linkOwner[n]; if (w) { r.who = w; return true; } return false; });
    const nm = pl2Name_(r.p, r.s, r.note); r.old.forEach(function (n) { map[n] = nm; });
  });
  return { list: list, map: map };
}
// Old layout -> new layout. Keeps a copy of both old tabs in prep-tabs-backup.json (Drive) first.
function pl2MigrateIfOld_(ss, out) {
  const pl = tabFind_(ss, PRL_TAB), wv = tabFind_(ss, WV_TAB);
  if (!pl || pl2Is_(pl) || pl.getLastRow() < 2) return false;
  out = out || [];
  const old = pcfgFromSheet_(ss);   // the old parser (PREP LIST + WEEK VIEW)
  const raw = pl.getRange(2, 1, pl.getLastRow() - 1, 6).getValues().map(function (r, i) { return { r: r, i: i, o: r[4] === '' || r[4] === null || isNaN(+r[4]) ? 1e9 + i : +r[4] }; })
    .sort(function (a, b) { return a.o - b.o || a.i - b.i; });
  const catalog = [], seen = {};
  raw.forEach(function (x) { const n = String(x.r[0] || '').trim().toUpperCase(), g = String(x.r[1] || '').trim().toUpperCase(); if (!n || !g || seen[n]) return; seen[n] = 1; catalog.push({ n: n, g: g, lbl: String(x.r[2] || '').trim().toUpperCase(), sl: Math.round(+x.r[3]) || 0, active: pcfgActive_(x.r[5]) }); });
  const cfg = old.cfg || { weekly: [[], [], [], [], [], [], []], every: [], lbl: {} };
  const tcfg = twConfig_(), names = (tcfg.employees || []).filter(function (e) { return !e.off; }).map(function (e) { return e.name; }), owner = {};
  (tcfg.tasks || []).forEach(function (t) { if (t.off) return; (t.l || []).forEach(function (n) { if (!owner[n]) owner[n] = t.rot ? t.rot[0] : t.emp; }); });
  const b = pl2Build_(catalog, cfg.weekly, cfg.every, cfg.lbl || {}, owner, names);
  try { logWriteJson_('prep-tabs-backup.json', { at: logStamp_(), prepList: pl.getDataRange().getValues(), weekView: wv ? wv.getDataRange().getValues() : [] }); } catch (x) { out.push('Backup NOT saved, PREP LIST left as it was: ' + x); return false; }
  pl2Write_(pl, b.list, names);
  // Collaborator tasks linked to a prep item keep working with the new names (column I of the task list).
  let relinked = 0;
  try {
    const ts = tabFind_(ss, TS_TAB);
    if (ts && ts.getLastRow() > 1) {
      const hd = ts.getRange(1, 1, 1, 20).getValues()[0].map(String), lc = hd.indexOf('Prep list item') + 1 || 14;
      const col = ts.getRange(2, lc, ts.getLastRow() - 1, 1).getValues();
      const nv = col.map(function (r) { const s = String(r[0] || ''); if (!s.trim()) return [r[0]]; const p = s.split(',').map(function (x) { const k = x.trim().toUpperCase(); return b.map[k] || x.trim(); }).join(', '); if (p !== s) relinked++; return [p]; });
      if (relinked) ts.getRange(2, lc, nv.length, 1).setValues(nv);
    }
  } catch (x) {}
  try { pl2RemapState_(b.map); } catch (x) {}
  out.push('PREP LIST changed to the new layout (' + b.list.length + ' rows; ' + catalog.length + ' prep items + the WEEK VIEW days). Old tabs saved in prep-tabs-backup.json.' + (relinked ? ' ' + relinked + ' task links renamed.' : ''));
  return true;
}
// Today's lists on the tablets keep what was already turned on / done: old item names -> new names.
function pl2RemapState_(map) {
  const doc = prepLoad_(); let ch = false;
  const mv = function (st, from, to) { ['n', 'x', 'u', 'e'].forEach(function (p) { if (st[p] && st[p][from] !== undefined) { if (st[p][to] === undefined) st[p][to] = st[p][from]; delete st[p][from]; ch = true; } }); };
  Object.keys(doc.l).forEach(function (loc) {
    const st = doc.l[loc];
    ['n', 'x', 'u', 'e'].forEach(function (p) { Object.keys(st[p] || {}).forEach(function (k) { if (k.indexOf('p_') === 0 && map[k.slice(2)] && map[k.slice(2)] !== k.slice(2)) mv(st, k, 'p_' + map[k.slice(2)]); }); });
    (st.c || []).slice().forEach(function (c) { if (!c || !map[c.n] || String(c.id).indexOf('dn_') === 0) return; mv(st, c.id, 'p_' + map[c.n]); st.c = st.c.filter(function (x) { return x !== c; }); ch = true; });
  });
  Object.keys(doc.pw || {}).forEach(function (d) { Object.keys(doc.pw[d]).forEach(function (n) { if (map[n] && map[n] !== n) { doc.pw[d][map[n]] = 1; delete doc.pw[d][n]; ch = true; } }); });
  if (ch) logWriteJson_(PREP_FILE, doc);
  return ch;
}
// WEEK VIEW = a picture of the PREP LIST days. Rewritten when it does not match (so typing there changes nothing).
function pl2View_(wv, cfg) {
  // only items with a WEEK VIEW category (PREP LIST column B) are shown, in that category
  const weekly = cfg.weekly.map(function (day) { return day.filter(function (w) { return w[2]; }).map(function (w) { return [w[2], w[1]]; }); });
  const every = (cfg.every || []).filter(function (x) { return x.wc; }).map(function (x) { return Object.assign({}, x, { c: x.wc }); });
  return wvWriteGrid_(wv, weekly, every, cfg.lbl || {}, [], true);
}
// One time: the old Stage column (B) becomes the WEEK VIEW category. "PICAR" + "ASADA" becomes the product "PICAR ASADA" (same name as before),
// and every item that already had days or an Every value gets its current Section as week view category, so the week view stays exactly as it is.
// A backup is saved first.
// One time: the PREP LIST "Shelf life (days)" column (P) is removed; the FIFO tab (column C) has the shelf life of the labels.
// A backup of the whole tab is saved first; if the backup fails nothing is changed.
function pl2DropShelf_(ss, notes) {
  const sh = tabFind_(ss, PRL_TAB); if (!sh || !pl2Is_(sh)) return false;
  if (!/^Shelf life/i.test(String(sh.getRange(1, 16).getValue() || '').trim()) || !/^Active/i.test(String(sh.getRange(1, 17).getValue() || '').trim())) return false;
  try { logWriteJson_('prep-shelf-backup-' + logNow_().date + '.json', { at: logStamp_(), values: sh.getDataRange().getValues() }); } catch (x) { if (notes) notes.push('PREP LIST shelf life column NOT removed (backup failed): ' + x); return false; }
  sh.deleteColumn(16);
  if (notes) notes.push('PREP LIST: Shelf life column removed (the FIFO tab has the shelf life). Backup saved.');
  return true;
}
function pl2StageMerge_(ss, notes) {
  const sh = tabFind_(ss, PRL_TAB); if (!sh || !pl2Is_(sh)) return false;
  if (!/^Stage/i.test(String(sh.getRange(1, 2).getValue() || '').trim())) return false;
  const n = Math.max(0, sh.getLastRow() - 1); if (!n) { sh.getRange(1, 2).setValue(PL2_HEAD[1]); return true; }
  const v = sh.getRange(2, 1, n, 17).getValues();
  try { logWriteJson_('prep-stage-backup-' + logNow_().date + '.json', { at: logStamp_(), values: v }); } catch (x) { if (notes) notes.push('PREP LIST stage column NOT changed (backup failed): ' + x); return false; }
  const out = v.map(function (r) {
    const p = String(r[0] || '').trim(), st = String(r[1] || '').trim().toUpperCase(), g = String(r[2] || '').trim().toUpperCase();
    if (!p && !st) return [r[0], r[1]];
    const any = r.slice(5, 12).some(alOnCell_) || !pl2Blank_(r[12]);
    const name = st && p.toUpperCase().indexOf(st + ' ') !== 0 ? st + ' ' + p : p;
    return [name, any && WV_CATORD.indexOf(g) >= 0 ? g : ''];
  });
  sh.getRange(2, 1, n, 2).setValues(out);
  sh.getRange(1, 2).setValue(PL2_HEAD[1]).setFontWeight('bold');
  try { sh.getRange(1, 1).setValue(PL2_HEAD[0]); sh.getRange(2, 2, Math.max(1, sh.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(WV_CATORD.slice(), true).setAllowInvalid(true).build()); } catch (x) {}
  if (notes) notes.push('PREP LIST: column B is now the WEEK VIEW category (stages joined to the product name, backup saved).');
  return true;
}


/* ===================== TASK LIST + PREP LIST: the owner's layout (Nov 3 2026) =====================
   Both tabs look the same: Who does it · ITEM · (How often | WEEK VIEW) · MON..SUN · (FIFO label) · Active · Status.
   A few values the tablets still need are kept in hidden columns at the far right (never typed by people):
     TASK LIST  M ID (history of each task) · N Start date (Every N days / One time) · O Prep list item (old links)
     PREP LIST  N Section on Today's task · O Every (days) · P Start date (for Every)
   The change is made once, automatically, after a backup of both tabs. Item names stay exactly the same on the tablets. */
const TS3_HEAD = ['Who does it', 'TASK ITEM', 'How often', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN', 'Active', 'Status (filled in automatically)', 'ID (do not change)', 'Start date', 'Prep list item'];
const PL3_HEAD = ['Who does it', 'PREP ITEM', 'WEEK VIEW', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN', 'FIFO label', 'Active', 'Status (filled in automatically)', "Section on Today's task", 'Every (days)', 'Start date (for Every)'];
const TS3 = { who: 0, task: 1, how: 2, d0: 3, act: 10, st: 11, id: 12, start: 13, link: 14 }, PL3 = { who: 0, item: 1, wc: 2, d0: 3, fifo: 10, act: 11, st: 12, sec: 13, ev: 14, an: 15 };
function ts3Is_(sh) { return !!sh && /^Who does it/i.test(String(sh.getRange(1, 1).getValue() || '').trim()) && /TASK ITEM/i.test(String(sh.getRange(1, 2).getValue() || '')); }
function pl3Is_(sh) { return !!sh && /^Who does it/i.test(String(sh.getRange(1, 1).getValue() || '').trim()) && String(sh.getRange(1, 4).getValue() || '').trim().toUpperCase() === 'MON'; }
function l3Names_() { try { return (twConfig_().employees || []).filter(function (e) { return !e.off; }).map(function (e) { return e.name; }); } catch (e) { return []; } }
function l3Dv_(list, strict) { return SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(!strict).build(); }
function ts3Format_(sh, n) {
  const rows = Math.max(1, n), R = function (c, k) { return sh.getRange(2, c, rows, k || 1); };
  try { sh.getRange(1, 1, Math.max(2, sh.getMaxRows()), Math.max(TS3_HEAD.length, sh.getMaxColumns())).clearDataValidations(); } catch (x) {}
  try { R(4, 7).insertCheckboxes(); } catch (x) {}
  try {
    sh.setFrozenRows(1); sh.setFrozenColumns(2);
    [150, 380, 140, 52, 52, 52, 52, 52, 52, 52, 70, 380, 110, 110, 170].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange(1, 1, 1, TS3_HEAD.length).setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
    sh.getRange(1, 4, rows + 1, 7).setHorizontalAlignment('center');
    R(1).setDataValidation(l3Dv_(l3Names_()));
    R(3).setDataValidation(l3Dv_(TS_HOW));
    R(11).setDataValidation(l3Dv_(['Yes', 'No'], true));
    R(14).setNumberFormat('m/d/yyyy');
  } catch (x) {}
  try { sh.hideColumns(13, 3); } catch (x) {}
}
function pl3Format_(sh, n) {
  const rows = Math.max(1, n), R = function (c, k) { return sh.getRange(2, c, rows, k || 1); };
  try { sh.getRange(1, 1, Math.max(2, sh.getMaxRows()), Math.max(PL3_HEAD.length, sh.getMaxColumns())).clearDataValidations(); } catch (x) {}
  try { R(4, 7).insertCheckboxes(); R(11).insertCheckboxes(); } catch (x) {}
  try {
    sh.setFrozenRows(1); sh.setFrozenColumns(2);
    [150, 300, 200, 52, 52, 52, 52, 52, 52, 52, 80, 70, 420, 190, 90, 120].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange(1, 1, 1, PL3_HEAD.length).setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
    sh.getRange(1, 4, rows + 1, 8).setHorizontalAlignment('center');
    R(1).setDataValidation(l3Dv_(l3Names_().concat(PL2_ROLES)));
    R(3).setDataValidation(l3Dv_(WV_CATORD.slice()));
    R(12).setDataValidation(l3Dv_(['Yes', 'No'], true));
    R(14).setDataValidation(l3Dv_(WV_CATORD.slice()));
    R(16).setNumberFormat('m/d/yyyy');
  } catch (x) {}
  try { sh.hideColumns(14, 3); } catch (x) {}
}
function l3Blank_(r, upto) { return r.slice(0, upto).every(function (v) { return v === '' || v === null || v === false || v === undefined; }); }
// One time: the current TASK LIST and PREP LIST are rewritten in the owner's layout (backup first; if the backup fails nothing changes)
function layout3_(ss, notes) {
  notes = notes || [];
  try { if (PropertiesService.getScriptProperties().getProperty('LAYOUT3_OFF') === '1') return false; } catch (e) {}   // safety switch: keep the previous layout
  const ts = tabFind_(ss, TS_TAB), pl = tabFind_(ss, PRL_TAB);
  let changed = false;
  if (ts && !ts3Is_(ts) && /^ID/i.test(String(ts.getRange(1, 1).getValue() || '')) && String(ts.getRange(1, 5).getValue() || '').trim().toUpperCase() === 'MON') {
    const n = Math.max(0, ts.getLastRow() - 1), v = n ? ts.getRange(2, 1, n, Math.max(15, Math.min(16, ts.getLastColumn()))).getValues() : [];
    try { logWriteJson_('tasklist-layout-backup-' + logNow_().date + '.json', { at: logStamp_(), head: ts.getRange(1, 1, 1, Math.max(1, ts.getLastColumn())).getValues()[0], values: v }); } catch (x) { notes.push('TASK LIST layout NOT changed (backup failed): ' + x); return false; }
    const rows = v.filter(function (r) { return !l3Blank_(r, 14); }).map(function (r) { return [r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[12] === '' ? 'Yes' : r[12], r[14] || '', r[0], r[11], r[13]]; });
    for (let i = 0; i < TS_SPARE; i++) rows.push(['', '', '', false, false, false, false, false, false, false, '', '', '', '', '']);
    try { ts.getBandings().forEach(function (b) { b.remove(); }); } catch (x) {}
    try { ts.showColumns(1, Math.max(TS3_HEAD.length, ts.getMaxColumns())); } catch (x) {}
    ts.clear();
    ts.getRange(1, 1, 1, TS3_HEAD.length).setValues([TS3_HEAD]);
    ts.getRange(2, 1, rows.length, TS3_HEAD.length).setValues(rows);
    ts3Format_(ts, rows.length);
    notes.push('TASK LIST: new layout (Who does it, TASK ITEM, How often, days, Active, Status). Backup saved.');
    changed = true;
  }
  if (pl && !pl3Is_(pl) && /^Product name/i.test(String(pl.getRange(1, 1).getValue() || '').trim()) && /^WEEK VIEW/i.test(String(pl.getRange(1, 2).getValue() || '').trim()) && String(pl.getRange(1, 6).getValue() || '').trim().toUpperCase() === 'MON' && /^Active/i.test(String(pl.getRange(1, 16).getValue() || '').trim())) {
    const n = Math.max(0, pl.getLastRow() - 1), v = n ? pl.getRange(2, 1, n, Math.max(17, Math.min(18, pl.getLastColumn()))).getValues() : [];
    try { logWriteJson_('preplist-layout-backup-' + logNow_().date + '.json', { at: logStamp_(), head: pl.getRange(1, 1, 1, Math.max(1, pl.getLastColumn())).getValues()[0], values: v }); } catch (x) { notes.push('PREP LIST layout NOT changed (backup failed): ' + x); return changed; }
    const rows = v.filter(function (r) { return !l3Blank_(r.slice(0, 5), 5) || r.slice(5, 12).some(alOnCell_) || !pl2Blank_(r[12]); }).map(function (r) {
      const item = [String(r[0] || '').trim(), String(r[4] || '').trim()].filter(Boolean).join(' ');   // product + note = the same name the tablets use
      return [r[3], item, r[1], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[14] === '' ? true : r[14], r[15] === '' ? 'Yes' : r[15], r[16] || '', r[2], r[12], r[13]];
    });
    for (let i = 0; i < PL2_SPARE; i++) rows.push(['', '', '', false, false, false, false, false, false, false, true, '', '', '', '', '']);
    try { pl.getBandings().forEach(function (b) { b.remove(); }); } catch (x) {}
    try { pl.showColumns(1, Math.max(PL3_HEAD.length, pl.getMaxColumns())); } catch (x) {}
    try { pl.getRange(1, 1, Math.max(1, pl.getMaxRows()), Math.max(PL3_HEAD.length, pl.getMaxColumns())).breakApart(); } catch (x) {}
    pl.clear();
    pl.getRange(1, 1, 1, PL3_HEAD.length).setValues([PL3_HEAD]);
    pl.getRange(2, 1, rows.length, PL3_HEAD.length).setValues(rows);
    pl3Format_(pl, rows.length);
    notes.push('PREP LIST: new layout (Who does it, PREP ITEM, WEEK VIEW, days, FIFO label, Active, Status). Backup saved.');
    changed = true;
  }
  if (changed) { try { tabLook_(ts); } catch (x) {} try { tabLook_(pl); } catch (x) {} try { hideStatus_(ss, notes); } catch (x) {} }
  return changed;
}
// TASK LIST row (new layout) -> the shape the task reader already understands [id, who, task, how, MON..SUN, start, active, link]
function ts3RowIn_(r) { return tsRowIn_([r[TS3.id], r[TS3.who], r[TS3.task], r[TS3.how]].concat(r.slice(TS3.d0, TS3.d0 + 7)).concat([r[TS3.start], r[TS3.act], r[TS3.link]])); }
// PREP LIST (new layout) -> the prep config. Section on Today's task = WEEK VIEW category, or the hidden section, or PREP GENERAL.
function pl3Parse_(sh, today) {
  const n = sh.getLastRow() - 1, rows = n > 0 ? sh.getRange(2, 1, n, PL3_HEAD.length).getValues() : [], st = [], items = [], weekly = [[], [], [], [], [], [], []], every = [], lbl = {}, sl = {}, seen = {};
  rows.forEach(function (r) {
    const days = r.slice(PL3.d0, PL3.d0 + 7).map(alOnCell_);
    if ([PL3.who, PL3.item, PL3.wc, PL3.ev].every(function (k) { return pl2Blank_(r[k]); }) && !days.some(Boolean)) { st.push(['']); return; }
    const nm = String(r[PL3.item] || '').trim().toUpperCase().replace(/\s+/g, ' ').slice(0, 60), who = String(r[PL3.who] || '').trim().toUpperCase().slice(0, 30);
    let wc = String(r[PL3.wc] || '').trim().toUpperCase().replace(/\s+/g, ' ').slice(0, 60), wcBad = false;
    if (wc && WV_CATORD.indexOf(wc) < 0) { wcBad = true; wc = ''; }
    const sec0 = String(r[PL3.sec] || '').trim().toUpperCase().slice(0, 60), g = wc || sec0 || 'PREP GENERAL', er = [];
    if (!nm) er.push('PREP ITEM is empty');
    if (nm && seen[nm]) er.push('This PREP ITEM is on another row');
    let ev = 0, an = '';
    if (!pl2Blank_(r[PL3.ev])) { ev = Math.round(+r[PL3.ev]); if (!(ev >= 2 && ev <= 30)) er.push('Every (days) must be 2 to 30'); an = tsDate_(r[PL3.an]); if (an === null) er.push('Start date not understood (use 11/3/2026)'); }
    if (er.length) { st.push(['⚠ ' + er.join('; ') + ' (not used)']); return; }
    seen[nm] = 1;
    if (!pcfgActive_(r[PL3.act])) { st.push(['Inactive (hidden on the tablets)']); return; }
    const fifo = pl2Blank_(r[PL3.fifo]) && r[PL3.fifo] !== false ? true : alOnCell_(r[PL3.fifo]);
    items.push({ n: nm, g: g, lbl: who, sl: 0, p: nm, s: '', f: fifo, wc: wc });
    if (who) lbl[nm] = who;
    days.forEach(function (on, k) { if (on) weekly[WV_COLS[k]].push([g, nm, wc]); });
    if (ev) every.push({ c: g, n: nm, e: ev, an: an || today, wc: wc });
    const when = days.map(function (on, k) { return on ? WV_HDR[k + 1].slice(0, 3) : ''; }).filter(Boolean).join(', ');
    st.push(['✓ ' + nm + (when ? ' · ' + when : '') + (ev ? ' · every ' + ev + ' days' : '') + (!when && !ev ? ' · any day' : '') + (wcBad ? ' · ⚠ WEEK VIEW not in the list (not on the week view)' : wc ? ' · week view: ' + wc : ' · not on the week view') + (fifo ? '' : ' · no FIFO label')]);
  });
  const cfg = { items: items, weekly: weekly, every: every, lbl: lbl, sl: sl };
  cfg.v = pcfgHash_(cfg);
  return { cfg: items.length ? cfg : null, st: st, err: items.length ? '' : 'The PREP LIST tab has no active items, so nothing was changed on the tablets.' };
}
// Every "Status (filled in automatically)" column is hidden (the script still fills it in). Right-click > Unhide to read one.
function hideStatus_(ss, out) {
  const done = [];
  Object.keys(TAB_NAMES).forEach(function (k) {
    try {
      const sh = tabFind_(ss, k); if (!sh) return;
      const w = Math.max(1, sh.getLastColumn()), head = sh.getRange(1, 1, Math.min(2, Math.max(1, sh.getLastRow())), w).getValues();
      for (let c = 0; c < w; c++) if (head.some(function (r) { return /^Status \(filled in automatically\)/i.test(String(r[c] || '').trim()); })) { sh.hideColumns(c + 1); done.push(sh.getName() + ' ' + String.fromCharCode(65 + c)); }
    } catch (e) {}
  });
  if (out && done.length) out.push('Status columns hidden: ' + done.join(', ') + '.');
  return done;
}
// "Who does it" dropdowns follow the COLLAB tab (refreshed when the names change)
function l3WhoDv_(ss) {
  const names = l3Names_(), key = 'who4:' + names.join(','), cache = CacheService.getScriptCache();
  try { if (cache.get('who3') === key) return; } catch (e) {}
  const ts = tabFind_(ss, TS_TAB), pl = tabFind_(ss, PRL_TAB);
  try { if (ts3Is_(ts)) ts.getRange(2, 1, Math.max(1, ts.getMaxRows() - 1), 1).setDataValidation(l3Dv_(names.concat(TW_ANY))); } catch (e) {}
  try { if (pl3Is_(pl)) pl.getRange(2, 1, Math.max(1, pl.getMaxRows() - 1), 1).setDataValidation(l3Dv_(names.concat(PL2_ROLES))); } catch (e) {}
  try { const al = tabFind_(ss, AL_TAB); if (al && alHas3_(al)) al.getRange(2, 14, Math.max(1, al.getMaxRows() - 1), 1).setDataValidation(l3Dv_(names)); } catch (e) {}
  try { cache.put('who3', key, 21600); } catch (e) {}
}

// Run ONCE from the editor: creates Tasks, Collaborators, PREP LIST, WEEK VIEW and FIFO HISTORY tabs (never overwrites a tab that has rows).
function setupSheets() { return setupTaskSheet(); }

/* ===================== FIFO LABELS history ===================== */
const RESTAURANT_ID = 'birriaz-san-jacinto';   // for future multi-restaurant use: every FIFO record carries its restaurant
const FIFO_TAB = 'FIFO HISTORY';
const FIFO_HEAD = ['Record ID', 'Restaurant', 'Created', 'Product', 'Custom item', 'Prepared', 'Used by', 'Shelf life (days)', 'Prepared by', 'Copies', 'Print status', 'Status updated', 'Device', 'Label type', 'Shelf life source'];
function fifoMdY_(ds) { const p = String(ds).split('-'); return p.length === 3 ? p[1] + '/' + p[2] + '/' + p[0] : ''; }
function fifoFind_(sh, id) { const n = sh.getLastRow(); if (n < 2) return -1; const v = sh.getRange(2, 1, n - 1, 1).getValues(); for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === id) return i + 2; return -1; }
function fifoLog_(r) {
  r = r || {}; const bad = function (m) { return { ok: false, fatal: true, error: 'invalid: ' + m }; };
  const id = String(r.id || ''); if (!/^[A-Za-z0-9_-]{6,40}$/.test(id)) return bad('id');
  const prod = String(r.product || '').trim().toUpperCase().slice(0, 60); if (!prod) return bad('product');
  const black = r.kind === 'black';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(r.prep || '') || (!black && !/^\d{4}-\d{2}-\d{2}$/.test(r.use || ''))) return bad('dates');
  const emp = String(r.emp || '').trim().toUpperCase().slice(0, 30); if (!emp) return bad('collaborator');
  const copies = Math.max(1, Math.min(50, Math.round(+r.copies || 1)));
  const sh = getTab_(FIFO_TAB, FIFO_HEAD);
  try { if (String(sh.getRange(1, FIFO_HEAD.length).getValue() || '') !== FIFO_HEAD[FIFO_HEAD.length - 1]) sh.getRange(1, 14, 1, 2).setValues([FIFO_HEAD.slice(13)]).setFontWeight('bold'); } catch (e) {}
  if (fifoFind_(sh, id) > 0) return { ok: true, dup: true };
  const src = black ? '' : ({ item: 'This product', category: 'Category default', none: 'NOT APPROVED - standard 6 days', prep: 'PREP LIST' }[r.slSrc] || '');
  sh.appendRow([id, NOVI_LOC + ' (' + RESTAURANT_ID + ')', String(r.at || logStamp_()).slice(0, 30), prod, black ? '' : r.custom ? 'YES - add it to the PREP LIST tab to keep it' : '', fifoMdY_(r.prep), black ? '' : fifoMdY_(r.use), black ? '' : (+r.sl || ''), emp, copies, String(r.status || 'Sent to printer').slice(0, 60), logStamp_(), String(r.dev || '').slice(0, 40), black ? 'Black sticker' : 'Regular', src]);
  return { ok: true };
}
function fifoStatus_(r) {
  r = r || {}; const id = String(r.id || '');
  if (!/^[A-Za-z0-9_-]{6,40}$/.test(id)) return { ok: false, fatal: true, error: 'invalid: id' };
  const sh = getTab_(FIFO_TAB, FIFO_HEAD), row = fifoFind_(sh, id);
  if (row < 0) return { ok: false, fatal: true, error: 'record not found' };
  sh.getRange(row, 11, 1, 2).setValues([[String(r.status || '').slice(0, 60), logStamp_()]]);
  return { ok: true };
}

/* ===================== INVENTORY items from the Google Sheet + reorder memory ===================== */
// The INVENTORY tab is the place to add, rename, hide or reorder inventory items (Item # + Product name).
// Tablets get changes through action=inv (checked when the Inventory page opens and every few minutes).
// Reorder memory = the Orders tab: every order email logs its items, and the average days between orders predicts the next one.
const INV_TAB = 'INVENTORY', INV_FILE = 'inv-config.json';
const INV_HEAD = ['Item #', 'Product name (shown in the app)', 'Industry name (supplier, used in order emails)', 'Category', 'Order', 'Active', 'Status (filled in automatically)'];
const INV_CATS = ['CHILES', 'SPICES', 'FOOD LIQUIDS', 'DRINKS', 'PEPSI', 'FRESH PRODUCE', 'FRUITS', 'CARRY OUT', 'DAIRY', 'MEATS', 'GRAINS/TORTILLA', 'PPE', 'CHEMICALS', 'SCRUBBERS', 'JANITORIAL', 'OFFICE', 'KITCHEN WARE', 'DESSERT'];
// Older INVENTORY tabs had no industry name column: insert it (column C) and copy the product names into it.
function invMigrate_(sh) {
  try { const h = String(sh.getRange(1, 3).getValue() || ''); if (!/^Category/i.test(h)) return false;
    sh.insertColumnAfter(2); sh.getRange(1, 3).setValue(INV_HEAD[2]).setFontWeight('bold'); sh.getRange(1, 2).setValue(INV_HEAD[1]); sh.getRange(1, 4).setValue(INV_HEAD[3]);
    const n = sh.getLastRow() - 1; if (n > 0) sh.getRange(2, 3, n, 1).setValues(sh.getRange(2, 2, n, 1).getValues());
    invCatRule_(sh); return true; } catch (x) { return false; } }
function invCatRule_(sh) { try { sh.getRange(2, 4, Math.max(1, sh.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(INV_CATS, true).setAllowInvalid(false).build()); } catch (x) {} }
// The INVENTORY tab, columns H to O, are filled in by the system (light red): on hand + reorder memory (was the dashboard).
const INV_AUTO = ['On hand', 'Last counted', 'Times ordered', 'Last order', 'Avg days between orders', 'Next order expected', 'Days left', 'Reorder status'];
function invHeadFix_(sh) {
  try {
    if (String(sh.getRange(1, 8).getValue() || '') === INV_AUTO[0] && String(sh.getRange(1, 15).getValue() || '') === INV_AUTO[7]) return false;
    sh.getRange(1, 8, 1, INV_AUTO.length).setValues([INV_AUTO]).setFontWeight('bold');
    try { [80, 150, 80, 130, 90, 140, 70, 190].forEach(function (w, i) { sh.setColumnWidth(8 + i, w); }); } catch (x) {}
    return true;
  } catch (x) { return false; }
}
// Who has how many: by item # (and by product name for older counts without an item #).
function cntIndex_() {
  const d = cntRead_(), code = {}, name = {};
  Object.keys(d.by).forEach(function (loc) {
    Object.keys(d.by[loc]).forEach(function (k) {
      const e = d.by[loc][k], c = String(e.c || '').trim().toUpperCase(), nm = String(k).trim().toUpperCase();
      [c ? [code, c] : null, [name, nm]].forEach(function (p) {
        if (!p) return; const o = p[0][p[1]] = p[0][p[1]] || { locs: {}, at: '' };
        o.locs[loc] = Number(e.n) || 0; if ((e.at || '') > o.at) o.at = e.at || '';
      });
    });
  });
  return { code: code, name: name };
}
function invTabRefresh_(countsOnly) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = tabFind_(ss, INV_TAB);
  if (!sh) return { ok: false, error: 'no INVENTORY tab' };
  invHeadFix_(sh);
  const n = sh.getLastRow() - 1; if (n < 1) return { ok: true, rows: 0 };
  const ab = sh.getRange(2, 1, n, 2).getValues(), tz = LOG_TZ, cn = cntIndex_();
  const ts = function (iso) { if (!iso) return ''; const d = new Date(iso); return isNaN(d) ? '' : Utilities.formatDate(d, tz, 'EEE MMM d, h:mm a'); };
  const key = function (r) { return [String(r[0] === null ? '' : r[0]).trim().toUpperCase(), String(r[1] || '').trim().toUpperCase()]; };
  const cv = ab.map(function (r) {
    const k = key(r); if (!k[0] && !k[1]) return ['', ''];
    const o = (k[0] && cn.code[k[0]]) || cn.name[k[1]]; if (!o) return ['', ''];
    const locs = Object.keys(o.locs);
    return [locs.length === 1 ? o.locs[locs[0]] : locs.map(function (l) { return l + ': ' + o.locs[l]; }).join(', '), ts(o.at)];
  });
  sh.getRange(2, 8, n, 2).setValues(cv);
  if (countsOnly) return { ok: true, rows: n };
  const today = dayNum_(new Date(), tz), stats = {};
  computeStats_(ordRows_().map(function (r) { return [r[0], 'ALL', String(r[2] || '').trim().toUpperCase(), r[3], r[4]]; }), today, function (d) { return dayNum_(new Date(d), tz); })
    .forEach(function (s) { if (s.code) stats[s.code] = s; });
  const fd = function (x) { return x === null || x === undefined ? '' : Utilities.formatDate(new Date(x * 86400000 + 12 * 3600000), 'UTC', 'EEE MMM d, yyyy'); };
  const sv = [], bg = [];
  ab.forEach(function (r) {
    const k = key(r), s = k[0] ? stats[k[0]] : null;
    if (!k[0] && !k[1]) { sv.push(['', '', '', '', '', '']); bg.push([LOCK_RED]); return; }
    if (!s) { sv.push([0, '', '', '', '', 'No orders yet']); bg.push(['#EEEEEE']); return; }
    sv.push([s.times, fd(s.last), s.avg === null ? '' : Math.round(s.avg * 10) / 10, fd(s.next), s.left === null ? '' : s.left, s.status]);
    bg.push([s.status.indexOf('OVERDUE') === 0 ? '#F4B6B6' : s.status.indexOf('ALMOST') === 0 ? '#FFE14D' : s.status === 'OK' ? '#CDEBCB' : '#E6E6E6']);
  });
  sh.getRange(2, 10, n, 6).setValues(sv);
  try { sh.getRange(2, 15, n, 1).setBackgrounds(bg); } catch (x) {}
  return { ok: true, rows: n };
}
// Kept under the old name: the menu, the daily 5 AM trigger and older setups call it.
function rebuildDashboard() { const out = []; invMergeTabs_(out); const r = invTabRefresh_(); Logger.log(out.concat([JSON.stringify(r)]).join('\n')); return r; }
function refreshDashboardDaily_() { rebuildDashboard(); }
function invDailyTrigger_(out) {
  try {
    if (ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'refreshDashboardDaily_'; })) return;
    ScriptApp.newTrigger('refreshDashboardDaily_').timeBased().everyDays(1).atHour(5).create();
    if (out) out.push('Daily 5 AM inventory refresh installed.');
  } catch (x) { if (out) out.push('Daily inventory refresh NOT installed: ' + x); }
}
// Run once from the editor (or setupSheets does it): installs the daily refresh and fills the numbers.
function setupDaily() { invDailyTrigger_(); return rebuildDashboard(); }

// Moves the old INVENTORY COUNTS / INVENTORY ORDERS tabs into the Drive files, then deletes them and the old dashboard.
// A tab is only deleted after its rows are safely in the file.
function invMergeTabs_(out) {
  out = out || [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = tabFind_(ss, 'Stock'), od = tabFind_(ss, 'Orders'), db = tabFind_(ss, 'Dashboard');
  if (!st && !od && !db) return out;
  try {
    if (st) {
      const d = cntRead_(); const n = cntMergeTab_(d); cntWrite_(d);
      const chk = logReadJson_(CNT_FILE);
      if (chk && chk.by) { ss.deleteSheet(st); out.push('INVENTORY COUNTS tab moved into ' + CNT_FILE + ' and removed' + (n ? ' (' + n + ' newer counts copied)' : '') + '.'); }
    }
  } catch (x) { out.push('INVENTORY COUNTS tab NOT moved: ' + x); }
  try {
    if (od) {
      const d = ordRead_(); const n = ordMergeTab_(d); logWriteJson_(ORD_FILE, d);
      const chk = logReadJson_(ORD_FILE), want = ordTabRows_().filter(function (r) { return r[0] && r[1] && r[3]; }).length;
      if (chk && Array.isArray(chk.rows) && chk.rows.length >= want) { ss.deleteSheet(od); out.push('INVENTORY ORDERS tab moved into ' + ORD_FILE + ' and removed (' + chk.rows.length + ' order lines kept' + (n ? ', ' + n + ' copied now' : '') + ').'); }
      else out.push('INVENTORY ORDERS tab kept: the copy did not check out.');
    }
  } catch (x) { out.push('INVENTORY ORDERS tab NOT moved: ' + x); }
  try { if (db) { ss.deleteSheet(db); out.push('INVENTORY DASHBOARD tab removed (its numbers are now on the INVENTORY tab, columns H to O).'); } } catch (x) { out.push('INVENTORY DASHBOARD tab NOT removed: ' + x); }
  return out;
}

// One time: categories and supplier (industry) names from the store's item # list, matched by Item #.
// Column B (the name shown in the app) is never changed on rows that already exist. Item #s with no industry name are not in this list.
const INV_FIX_KEY = 'INV_FIX_1';
const INV_CAT_OLD = { 'LIQUID FOOD': 'FOOD LIQUIDS', 'GRAINS/BREAD/TORTILLAS': 'GRAINS/TORTILLA', 'DESSERTS': 'DESSERT' };
function invApplyFix_(sh, out) {
  const P = PropertiesService.getScriptProperties();
  if (P.getProperty(INV_FIX_KEY)) return false;
  invMigrate_(sh); invHeadFix_(sh);
  const fix = {}, ord = [];
  INV_FIX.split('\n').forEach(function (l) { const p = l.split('|'); if (p.length < 3) return; const c = p[0].trim().toUpperCase(); if (!c || fix[c]) return; fix[c] = { cat: p[1].trim().toUpperCase(), ind: p.slice(2).join('|').trim() }; ord.push(c); });
  const n = sh.getLastRow() - 1, rows = n > 0 ? sh.getRange(2, 1, n, 6).getValues() : [], have = {};
  let upd = 0, maxO = 0;
  const cd = rows.map(function (r) {
    const c = String(r[0] === null ? '' : r[0]).trim().toUpperCase(); if (c) have[c] = 1;
    if (r[4] !== '' && r[4] !== null && !isNaN(+r[4])) maxO = Math.max(maxO, +r[4]);
    let ind = r[2], cat = String(r[3] || '').trim().toUpperCase();
    if (INV_CAT_OLD[cat]) cat = INV_CAT_OLD[cat];
    const f = fix[c]; if (f) { ind = f.ind; cat = f.cat; upd++; }
    return [ind, cat];
  });
  try { sh.getRange(2, 4, Math.max(1, sh.getMaxRows() - 1), 1).clearDataValidations(); } catch (x) {}   // the old category list would reject the new names
  if (n > 0) sh.getRange(2, 3, n, 2).setValues(cd);
  const add = [];
  ord.forEach(function (c) { if (have[c]) return; maxO += 10; add.push([c, fix[c].ind.toUpperCase(), fix[c].ind, fix[c].cat, maxO, 'Yes', '']); });
  if (add.length) { const at = sh.getLastRow() + 1; try { sh.getRange(at, 1, add.length, 1).setNumberFormat('@'); } catch (x) {} sh.getRange(at, 1, add.length, 7).setValues(add); }
  invCatRule_(sh);
  P.setProperty(INV_FIX_KEY, logStamp_());
  if (out) out.push('INVENTORY: industry names + categories filled in for ' + upd + ' items (column B left as it was); ' + add.length + ' new items added from the item # list.');
  return true;
}
const INV_FIX = `1|CHILES|CHILE DE ARBOL TOSTADO #5
2|CHILES|CHIPOTLE ADOVO Canned #10
3|CHILES|GUAJILLO STEAMLESS #25
4|CHILES|COSTENO
5|CHILES|PASILLA / ANCHO STEAMLESS #25
12|SPICES|ALL SPICE WHOLE
13|SPICES|BAY LEAF WHOLE
14|SPICES|CLOVES WHOLE
15|SPICES|MEJORANA WHOLE
16|SPICES|OREGANO WHOLE
17|SPICES|THYME WHOLE
21|SPICES|ACHIOTE PASTE
22|SPICES|PAPRIKA POWDER #5
25|SPICES|SALT PACKS INDIVIDUAL
26|SPICES|SALT
27|SPICES|GARLIC SALT
28|SPICES|CHICKEN SEASONING CHEF MERITO 28LBS
29|SPICES|FAJITA SEASONING CHEF MERITO
30|SPICES|MEAT Seasoning Chef Merito 28 LBS
31|SPICES|PEPPER GROUND 18 MESH
32|SPICES|GARLIC GRANULATED
33|SPICES|CUMIN POWDER
34|SPICES|MEAT TENDERRIZER CHEF'S QUALITY
35|SPICES|ONION GRANULATED
36|SPICES|FISH SEASONING Chef Merito
37|SPICES|CALDO DE CAMARON KNORR 4.4lbs
38|SPICES|OLD BAY SEASONING
40|SPICES|BEEF BOUILLON KNORR 4.4lbs
41|SPICES|CHICKEN BOUILLON KNORR 4.4lbs
42|SPICES|TOMATO BOUILLON KNORR 4.4lbs
43|SPICES|ALPASTOR seasoning Chef Merito
49|SPICES|CINNAMON STICKS
50|SPICES|CHOCOLATE ABUELITA
51|SPICES|SUGAR GRANULATED 25LBS C&H
52|SPICES|PILONCILLO
53|SPICES|HOT CHEETOS PUFFS
54|SPICES|SUGAR SATIVIA
55|SPICES|INDIVIDUAL SUGAR DISPENSER
60|FOOD LIQUIDS|CHIPOTLE MAYO Ken's sauces
61|FOOD LIQUIDS|MEAT MARINADE chef Merito
62|FOOD LIQUIDS|CHICKEN seasoning chef Merito
63|FOOD LIQUIDS|CLEAR FRYING OIL
64|FOOD LIQUIDS|OLIVE OIL
65|FOOD LIQUIDS|WORCESTERSHIRE
66|FOOD LIQUIDS|VINEGAR WHITE
67|FOOD LIQUIDS|LIME JUICE
68|FOOD LIQUIDS|SUNNY D
69|FOOD LIQUIDS|NACHO CHEESE
70|FOOD LIQUIDS|AGED WHITE CHEESE SAUCE
71|FOOD LIQUIDS|YELLOW CHEESE SAUCE
72|FOOD LIQUIDS|QUESO BLANCO
73|FOOD LIQUIDS|CHILE CON QUESO
74|FOOD LIQUIDS|EVAPORATED MILK
75|FOOD LIQUIDS|LECHERA
76|FOOD LIQUIDS|PINEAPPLE JUICE
77|FOOD LIQUIDS|LIQUID BUTTER
78|FOOD LIQUIDS|VANILLA EXTRACT
80|DRINKS|HORCHATA SYRUP Felbro
81|DRINKS|JAMAICA SYRUP
82|DRINKS|LEMONADE SYRUP Felbro
83|DRINKS|TAMARINDO SYRUP
84|DRINKS|PA LEMONADE
85|DRINKS|PA HORCHATA
86|DRINKS|PA HIBISCUS
87|DRINKS|PA MANGO
88|DRINKS|PA PINEAPPLE
89|DRINKS|PA TAMARIND
90|DRINKS|PA WATERMELON
91|DRINKS|PINEAPPLE POWDER
92|DRINKS|CANTELOPE POWDER
93|DRINKS|COKE CAN
94|DRINKS|COKE ZERO CAN
95|DRINKS|DIET COKE CAN
96|DRINKS|DR PEPPER CAN
97|DRINKS|PEPSI CAN
98|DRINKS|DIET PEPSI CAN
99|DRINKS|SPRITE CAN
100|DRINKS|SQUIRT CAN
101|DRINKS|FANTA STRAWBERRY CAN
102|DRINKS|FANTA ORANGE CAN
103|DRINKS|FANTA PINEAPPLE CAN
104|DRINKS|Poppi Punch pop
105|DRINKS|Poppi STRAWBERRY LEMON
106|DRINKS|Poppi WILD BERRY
107|DRINKS|Poppi RASPBERRY ROSE
110|DRINKS|PINK LEMONADE SYRUP Felbro
112|DRINKS|GLASS PEPSI 16.9
113|DRINKS|SPRITE GLASS BOTTLE
114|DRINKS|GLASS COKE BOTTLE 16.9
115|DRINKS|AQUAFINA
116|DRINKS|TOPO CHICO
117|DRINKS|JARITOS MANDARIN
118|DRINKS|JARITOS STRAWBERRY
119|DRINKS|JARITOS LIME
120|DRINKS|JARITOS TAMARIND
121|DRINKS|JARITOS GUAVA
122|DRINKS|JARITOS PINEAPPLE
123|DRINKS|JARITOS GRAPEFRUIT
124|FRESH PRODUCE|POTATO
127|FRESH PRODUCE|FRESH STRAWBERRY
128|FRESH PRODUCE|FAJITAS
129|FRESH PRODUCE|RED BELL PEPPERS
130|FRESH PRODUCE|ORANGE/YELLOW BELL PEPPER
131|FRESH PRODUCE|GREEN BELL PEPPER
133|FRESH PRODUCE|WHITE ONION MED #50
134|FRESH PRODUCE|WHITE ONION Jumbo #25
135|FRESH PRODUCE|WHITE ONION DICED 1/4
136|FRESH PRODUCE|WHITE ONION DICED 3/8
137|FRESH PRODUCE|RED ONION Jumbo #50
138|FRESH PRODUCE|RED ONION Jumbo #25
139|FRESH PRODUCE|PURPLE ONION SLICED 1/4
142|FRESH PRODUCE|TOMATOE DICED
143|FRESH PRODUCE|ROMAINE LETTUCE
144|FRESH PRODUCE|CABBAGE GREEN
145|FRESH PRODUCE|CABBAGE PURPLE
147|FRESH PRODUCE|CORN FIRE ROASTED FROZEN
148|FRESH PRODUCE|AVOCADO 60
149|FRESH PRODUCE|ROMA TOMATOES
150|FRESH PRODUCE|TOMATILLO Crushed CANNED #10
151|FRESH PRODUCE|LEMON 140ct
152|FRESH PRODUCE|CILANTRO
153|FRESH PRODUCE|GARLIC PEELED
154|FRESH PRODUCE|GINGER fresh
155|FRESH PRODUCE|BEETS POWDER
156|FRESH PRODUCE|TOMATO WHOLE PEELED CANNED #10
167|FRESH PRODUCE|SERRANO
168|FRESH PRODUCE|HABANERO
169|FRESH PRODUCE|JALAPENO
170|FRUITS|PINEAPPLE
171|FRUITS|WATERMELON
172|FRUITS|CANTELOPE
173|FRUITS|CUCUMBER
174|FRUITS|FROZEN STRAWBERRY
175|FRUITS|JAMAICA FLOWER
176|FRUITS|BLUE BERRIES
177|FRUITS|RASBERRY
178|FRUITS|DRAGON FRUIT
179|FRUITS|FROZEN PINNEAPLE
180|CARRY OUT|9X9 BLK 1-COMP HINGED CONTAINER
181|CARRY OUT|8x8 Ecopac Black/clear container 150 ct
183|CARRY OUT|37 OZ ROUND BOWL(RAMEN) microwave
184|CARRY OUT|8X8 HINGED 3 COMP white
185|CARRY OUT|bowl, fiber 32oz burrito
186|CARRY OUT|lid, dome 32 oz bowl
188|CARRY OUT|2 OZ FP-P200-PP KARAT
189|CARRY OUT|2 OZ FP-PL200-PET KARAT
191|CARRY OUT|6 OZ CONT PAPER C-KDP6W
192|CARRY OUT|6 OZ LIDS C-KDL96-PP
194|CARRY OUT|16 OZ PAPER C-KDP16W
195|CARRY OUT|16 OZ LIDS C-KDL112-PP
196|CARRY OUT|5 OZ PAPER C-KDP5 PINK
197|CARRY OUT|5 OZ LIDS C-KDL87-PP
199|CARRY OUT|16 OZ CLEAR CUP C-KC16
200|CARRY OUT|24 OZ CLEAR CUP C-KC24
201|CARRY OUT|12-24 OZ CLEAR LID STRAW C-KC626TS
202|CARRY OUT|32 OZ CLEAR CUP C-KC32
203|CARRY OUT|32 OZ CLEAR LID STRAW C-KC107TS
205|CARRY OUT|9-1/4 X 7-1/4 BLACK FOAM #4S
206|CARRY OUT|5-1/4 X 5-1/4 BLACK FOAM #1S
207|CARRY OUT|PAPER PLATE #200
209|CARRY OUT|FOIL ROLL 12"
210|CARRY OUT|FOIL SHEETS 9 X 10.75
211|CARRY OUT|INSULATED FOIL WRAP 14X16
212|CARRY OUT|12"X12" CHECKER DELI PAPER BLK
214|CARRY OUT|FORK HEAVY BLK
215|CARRY OUT|SOUP SPOON MEDIUM ( WHITE)
216|CARRY OUT|SOUP SPOON MEDIUM (WHITE)
217|CARRY OUT|KNIFE MEDIUM (BLK)
219|CARRY OUT|Sandwish Zip loc bags
220|CARRY OUT|CARRY OUT BAGS T SHIRT
221|CARRY OUT|ZIPLOC BAGS 1 GAL
222|CARRY OUT|TOOTH PICKS MINT
223|CARRY OUT|CUP HOLDER DUAL
224|CARRY OUT|CUP HOLDER 4
225|CARRY OUT|STRAWS 10.25"
227|CARRY OUT|ALUMINUM HALF PAN SHALLOW
228|CARRY OUT|ALUMINUM HALF PAN LID
229|CARRY OUT|CATERING BOX 21X13X3.5
230|CARRY OUT|TAMPER SAFE 1X3 SEALED LABEL
231|CARRY OUT|TAMPER SAFE 8X1 FLAT LID DRINK LABEL
232|CARRY OUT|MONDAY STICKER
233|CARRY OUT|TUESDAY STICKER
234|CARRY OUT|WEDNESDAY STICKER
235|CARRY OUT|THURSDAY STICKER
236|CARRY OUT|FRIDAY STICKER
237|CARRY OUT|SATURDAY STICKER
238|CARRY OUT|SUNDAY STICKER
239|CARRY OUT|USED FIRST STICKER
240|CARRY OUT|PRODUCT LABEL
241|CARRY OUT|6" DESSERT HING CONT
243|CARRY OUT|BAGS FOR BIRRIA
244|CARRY OUT|OVAL CONTAINER BLACK WHIT LIDS 32-OZ 1/150 BOX
245|CARRY OUT|BAGS FOR RICE AND BEANS
247|CARRY OUT|water cup trans y7
249|CARRY OUT|PIZZA BOX 14" X 3"
250|DAIRY|MONTEREY JACK CHEESE
251|DAIRY|MEXICAN MIX CHEESE
252|DAIRY|MEXICAN CREMA
253|DAIRY|SOUR CREAM
254|DAIRY|MAYO
260|MEATS|BIRRIA MEAT
261|MEATS|CHICKEN LEG
262|MEATS|ASADA MEAT
263|MEATS|BOONLESS PORK SHOULDER BUTT
264|MEATS|SHRIMP
268|GRAINS/TORTILLA|MAZA GUAJILLO
269|GRAINS/TORTILLA|MAZA REG
270|GRAINS/TORTILLA|TOP RAMEN
271|GRAINS/TORTILLA|TAPATIO SPICY RAMEN 3.8 OZ BOWL
273|GRAINS/TORTILLA|14" FLOUR EL METATE
274|GRAINS/TORTILLA|10" FLOUR TORTILLAS (QUESADILLA)
275|GRAINS/TORTILLA|5" TORTILLA FLOUR
276|GRAINS/TORTILLA|MULITAS COMAL NEJA
277|GRAINS/TORTILLA|5" TORTILLA CORN (CALIDAD)
278|GRAINS/TORTILLA|5" YELLOW EL METATE (alpastor)
279|GRAINS/TORTILLA|4.5" TORTILLA TAQUERA
281|GRAINS/TORTILLA|BURGER BREAD (SMASH BURGER)
282|GRAINS/TORTILLA|BURGER BREAD (BIRRIA)
283|GRAINS/TORTILLA|TORTA BREAD( TELERA) CHAPALITA
284|GRAINS/TORTILLA|BEANS PINTO TRIPPLE WASHED #50
285|GRAINS/TORTILLA|RICE, JAZMINE #50
286|GRAINS/TORTILLA|3/8 STRAIGHT FRIES
287|GRAINS/TORTILLA|GRILL CHEESE BREAD
289|GRAINS/TORTILLA|HOT CHEETOS
290|PPE|GLOVES SIZE S (BLK)
291|PPE|GLOVES SIZE M (BLK)
292|PPE|GLOVES SIZE L (BLK)
293|PPE|GLOVES SIZE XL (BLK)
294|PPE|SINGLE USE GLOVES SIZE M
295|PPE|HAIRNET
296|PPE|PLASTIC APRIN
297|PPE|HAND CREAM
299|PPE|hybrid large gloves
300|PPE|hybrid medium gloves
301|PPE|hybrid xl gloves
303|PPE|ELARA EASYFIT MAX POLY GLOVES L
304|PPE|ELARA EASYFIT MAX POLY GLOVES M
305|PPE|VINYL GLOVES WITH OUT POWDER S
306|PPE|VINYL GLOVES WITH OUT POWDER M
307|PPE|VINYL GLOVES WITH OUT POWDER L
308|PPE|VINYL GLOVES WITH OUT POWDER XL
310|CHEMICALS|DISINFECTANT SANITIZER DS-100
312|CHEMICALS|DISH SOAP
314|CHEMICALS|HEAVY DUTY DEGREASER (RED)
316|CHEMICALS|DEGREASER (PURPLE)
318|CHEMICALS|PINE CLEANER MOPPING
319|CHEMICALS|FRYER CLEANER
320|CHEMICALS|GRILL DEGREASER
321|CHEMICALS|GRIDDLE CLEANER
322|CHEMICALS|HAND SOAP PINK LIQUID 1GAL
323|CHEMICALS|DIAL FOAMING HAND WASH DIA32100
324|CHEMICALS|WINDEX
325|CHEMICALS|BLEACH
326|CHEMICALS|TOILET BOWL SOAP
327|CHEMICALS|CLOROX WIPES
328|CHEMICALS|LAUNDRY SOAP
329|CHEMICALS|BLEACH
340|SCRUBBERS|GREEN SCRUBBERS
341|SCRUBBERS|STAINLESS STEEL SCRUBBER
342|SCRUBBERS|HEAVY DUTY POT SCRUBBERS
343|SCRUBBERS|GRILL SCREEN
344|SCRUBBERS|GRILL BRICKS
345|SCRUBBERS|MOPHEAD
346|SCRUBBERS|TOILET SCRUBBER
355|JANITORIAL|DINNER NAPKINS INTERFOLD 6.5X8.5(13X8.5)
356|JANITORIAL|DINNER NAPKINS INTERFOLD 6.5X8.5(13X8.5)
357|JANITORIAL|MULTI FOLD TOWELS
358|JANITORIAL|karat paper roll JS-RTK750
359|JANITORIAL|"TORK" TISSUE, BATH 2PLY 7" JUMBO
360|JANITORIAL|TOILET SEAT COVER
361|JANITORIAL|TRASH CAN (HAND SINK)
362|JANITORIAL|TRASH CAN PREP
363|JANITORIAL|SLIM TRASH CAN
364|JANITORIAL|36 QUART MOP BUCKET- BALDE PARA TRAPEADOR
365|JANITORIAL|TRASH BAG WHITE 7-10 GAL
366|JANITORIAL|TRASH BAG 32 GALLON
367|JANITORIAL|KITCHEN MOP HANDLE
368|JANITORIAL|DINNING MOP HANDLE
369|JANITORIAL|BROOM
370|JANITORIAL|DUST PAN
371|JANITORIAL|PLUNGER
372|JANITORIAL|KITCHEN TOWELS
373|JANITORIAL|GLAD AIR SPRAYER
374|JANITORIAL|BATHROOM AIR FRESHENER
375|JANITORIAL|DINING AUTO SPRAY FRESHENER
386|OFFICE|END OF SHIFT CASH ENVELOPE
387|OFFICE|TIPS ENVELOPE
388|OFFICE|RECEIPT PRINTER PAPER ROLL
389|OFFICE|COUNTER TAPE-
390|OFFICE|DRY ERASE MARKERS
391|OFFICE|PERMANENT MARKERS
392|OFFICE|HIGHLIGHTER
393|OFFICE|COUNTERFEIT MARKER
394|OFFICE|INK PEN
395|OFFICE|$0.01
396|OFFICE|$0.05
397|OFFICE|$0.10
398|OFFICE|PRINTER PAPER
399|OFFICE|LAMINATING PAPER
406|PEPSI|STRAWBERRY WATERMELON SLUSHI
407|PEPSI|MANGO SLUSHI
408|PEPSI|PEPSI BIB
409|PEPSI|DIET PEPSI BIB
410|PEPSI|STARRY BIB
411|PEPSI|TROPICANA PINK LEM BIB
412|PEPSI|BRISK RASBERRY BIB
413|PEPSI|LIPTON UN ICE TEA BIB
414|PEPSI|CRUSH BIB
415|PEPSI|DR PEPPER BIB
416|KITCHEN WARE|KITCHEN LIGHTER
417|KITCHEN WARE|PAN LINERS BAGS
418|KITCHEN WARE|THERMOMETER
419|KITCHEN WARE|BUTANE FUEL TORCH
425|KITCHEN WARE|CAMBRO - 6-QT ROUND CONTAINER
427|KITCHEN WARE|CAMBRO -STORAGE CONTAINER, 12" X 18" X 9" CLEAR
430|KITCHEN WARE|CAMWEAR DRAIN SHELF, 1/6 SIZE
431|KITCHEN WARE|CAMWEAR DRAIN SHELF, 1/3 SIZE
432|KITCHEN WARE|CAMWEAR DRAIN SHELF, FULL SIZE
439|DESSERT|JELLO
440|DESSERT|RICE PUDDING
441|DESSERT|FLAN
442|DESSERT|PAND DE ELOTE
443|DESSERT|CHEESECAKE CASERO
444|DESSERT|NY CHEESE CAKE
445|DESSERT|FRESAS CON CREMA
446|DESSERT|CARROT CAKE`;
const INV_SEED = `209|FOIL PAPER
227|FAMILY ALUMINUM TRAY
228|FAMILY ALUMINUM LID
296|PLASTIC APRON
211|BURRITO FOIL PAPER
417|PLASTIC PAN LINERS
221|ZIPLOC BAGS 1 GAL
345|MOP HEAD
223|CUP HOLDERS
277|TORTILLAS QUESABIRRIAS
387|TIPS ENVELOPES
386|END OF SHIFT ENVELOPE
398|PRINTER PAPER
399|LAMINATING PAPER
389|TAPE (DESK)
391|PERMANENT MARKERS
390|BLACK MARKERS (DRY ERASERS)
394|PENS
26|SALT
284|BEANS
285|RICE
282|BURGER BREAD
283|TORTA BREAD
287|GRILL CHEESE BREAD
110|PINK LEMONADE
241|DESSERT CONTAINER
249|PIZZA BOX
184|PLATES FOR SPECIALS
206|VEGGIE PLATE FOAM
205|LARGE FOAM PLATE
80|HORCHATA
82|LEMONADE REGULAR
322|HAND SOAP
318|PINESOL
328|LAUNDRY SOAP
329|BLEACH
319|FRYER CLEANER
321|GRIDDLE LIQUID CLEANER
324|WINDEX
326|TOILET BOWL SOAP
323|BATHROOM HAND SOAP
365|BATHROOM TRASH BAG
366|TRASH BAG
342|DISH SCRUBBERS
344|GRILL BRICKS
316|DEGREASER
310|DISINFECTANT SANITIZER
312|DISH SOAP
147|CORN
174|FROZEN STRAWBERRY
229|CATERING BOX
262|BIRRIA/ASADA MEAT
179|PINEAPPLE
261|CHICKEN
263|ALPASTOR
63|OIL
295|HAIRNET
373|GLAD AIR FRESHNER
180|CONTAINER QUESABIRRIAS
183|RAMEN BOWLS
207|PAPER PLATE (RE HEAT)
217|KNIFE
216|SOUP SPOON
214|FORK
307|BLACK GLOVES LARGE
304|GLOVES MEDIUM
303|GLOVES LARGE
225|STRAWS
192|CONSOME LID
201|DRINK LIDS MEDIUM
203|DRINK LIDS LARGE
191|CONSOME CUP
199|DRINK SMALL CUP
200|DRINK MEDIUM CUP
202|DRINK LARGE CUP
189|SALSA CUP
188|SALSA LID
220|CARRY OUT BAGS
232|MONDAY STICKER
233|TUESDAY STICKER
234|WEDNESDAY STICKER
235|THURSDAY STICKER
236|FRIDAY STICKER
237|SATURDAY STICKER
238|SUNDAY STICKER
239|USED FIRST STICKER
240|PRODUCT LABEL
419|FUEL TORCH
212|CHECKER PAPER
30|MEAT SALT
12|ALL SPICE WHOLE
14|CLOVES WHOLE
219|Sandwish Zip loc bags
29|FAJITA SEASONING CHEF MERITO
27|GARLIC SALT
32|GARLIC GRANULATED
35|ONION GRANULATED
34|MEAT TENDERRIZER
155|BEETS POWDER
60|CHIPOTLE MAYO
69|NACHO CHEESE
76|PINEAPPLE JUICE
43|ALPASTOR seasoning
2|CHIPOTLE ADOVO
22|PAPRIKA POWDER
40|BEEF BOUILLON
42|TOMATO BOUILLON
21|ACHIOTE PASTE
36|FISH SEASONING Chef Merito
37|CALDO DE CAMARON KNORR 4.4lbs
38|OLD BAY SEASONING
243|BAGS FOR BIRRIA
245|BAGS FOR BIRRIA
17|THYME WHOLE
16|OREGANO WHOLE
15|MEJORANA WHOLE
33|CUMIN POWDER
31|PEPPER GROUND 18 MESH
13|BAY LEAF WHOLE
65|WORCESTERSHIRE
66|VINEGAR WHITE
41|CHICKEN BOUILLON
150|TOMATILLO CANNED
156|TOMATO CANNED
133|WHITE ONION
138|RED ONION Jumbo
149|ROMA TOMATOES
151|LEMONS
269|MAZA REG
268|MAZA GUAJILLO
124|POTATO
67|LIME JUICE
68|SUNNY D
153|GARLIC PEELED
154|GINGER fresh
250|MONTEREY JACK CHEESE
148|AVOCADO 60
254|MAYO
61|MEAT MARINADE LIQUID
168|HABANERO
128|FAJITAS
143|ROMAINE LETTUCE
167|SERRANO
169|JALAPENO
144|CABBAGE GREEN
145|CABBAGE PURPLE
152|CILANTRO
1|CHILE DE ARBOL
3|GUAJILLO
5|PASILLA / ANCHO
49|CINNAMON STICKS
51|SUGAR
78|VANILLA EXTRACT
357|MULTI FOLD TOWELS
360|TOILET SEAT COVER
359|TOILET PAPER
355|DINNER NAPKINS
358|HAND ROLL PAPER
195|LARGE CONSOME LID
273|BURRITO TORTILLA
274|QUESADILLA TORTILLA
210|FOIL SHEETS
194|LARGE CONSOME CUP
439|JELLO
440|RICE PUDDING
441|FLAN
442|PAND DE ELOTE
445|FRESAS CON CREMA
446|CARROT CAKE
444|NY CHEESE CAKE
74|EVAPORATED MILK
50|CHOCOLATE ABUELITA
170|PINEAPPLE
173|CUCUMBER
252|MEXICAN CREMA
127|FRESH STRAWBERRY
89|PA TAMARIND
90|PA WATERMELON
86|PA HIBISCUS
87|PA MANGO
93|COKE CAN
94|COKE ZERO CAN
95|DIET COKE CAN
114|COKE MEDIO LITRO
112|PEPSI GLASS
271|RAMEN
222|TOOTH PICKS MINT
25|SALT PACKS
54|SUGAR SATIVIA
55|INDIVIDUAL SUGAR DISPENSER
230|SEALED BAG LABELS
231|SEALED DRINK LABELS
388|RECEIPT PRINTER PAPER
408|PEPSI BIB
409|DIET PEPSI BIB
410|STARRY BIB
411|TROPICANA PINK LEM BIB
406|STRAWBERRY WATERMELON SLUSHI
412|BRISK RASBERRY BIB
413|LIPTON UN ICE TEA BIB
414|CRUSH BIB
415|DR PEPPER BIB
407|MANGO SLUSHI`;
function invFromSheet_(ss) {
  const sh = tabFind_(ss, INV_TAB);
  if (!sh) return { err: 'no tab' };
  invMigrate_(sh);
  const n = sh.getLastRow() - 1, rows = n > 0 ? sh.getRange(2, 1, n, 7).getValues() : [], st = [], seen = {}, items = [];
  const list = rows.map(function (r, i) { return { r: r, i: i, o: r[4] === '' || r[4] === null || isNaN(+r[4]) ? 1e9 + i : +r[4] }; });
  list.forEach(function (x) { st[x.i] = [String(x.r[6] || '')]; });
  list.slice().sort(function (a, b) { return a.o - b.o || a.i - b.i; }).forEach(function (x) {
    const r = x.r, c = String(r[0] === null ? '' : r[0]).trim().toUpperCase().slice(0, 12), nm = String(r[1] || '').trim().toUpperCase().slice(0, 60), ind = String(r[2] || '').trim().toUpperCase().slice(0, 80), cat = String(r[3] || '').trim().toUpperCase().slice(0, 40);
    if (!c && !nm) { st[x.i] = ['']; return; }
    if (!c) { if (!/^Requested/.test(st[x.i][0])) st[x.i] = ['⚠ Item # is empty (not used)']; return; }
    if (!/^[A-Z0-9-]{1,12}$/.test(c)) { st[x.i] = ['⚠ Item # can only have letters, numbers or - (not used)']; return; }
    if (!nm) { st[x.i] = ['⚠ Product name is empty (not used)']; return; }
    if (seen[c]) { st[x.i] = ['⚠ Item # ' + c + ' is used twice (not used)']; return; }
    seen[c] = 1;
    if (!pcfgActive_(r[5])) { st[x.i] = ['Inactive (hidden on the tablets)']; return; }
    items.push({ c: c, n: nm, ind: ind || nm, cat: cat }); st[x.i] = ['✓ OK'];
  });
  if (!items.length) return { err: 'The INVENTORY tab has no active items, so nothing was changed on the tablets.', status: st };
  const cfg = { items: items }; cfg.v = pcfgHash_(cfg);
  return { cfg: cfg, status: st };
}
function invSync_(ss) {
  const pre = [];
  try { const s0 = tabFind_(ss, INV_TAB); if (s0) invApplyFix_(s0, pre); } catch (x) { pre.push('Item # list NOT applied: ' + x); }
  try { invMergeTabs_(pre); } catch (x) {}
  try { if (!PropertiesService.getScriptProperties().getProperty('INV_AUTO_TRG')) invAutoTrigger_(pre); } catch (x) {}
  const r = invFromSheet_(ss);
  if (r.err === 'no tab') return { ok: false, error: 'no INVENTORY tab yet (run setupSheets once)' };
  try {
    const sh = tabFind_(ss, INV_TAB);
    if (r.status && r.status.length) sh.getRange(2, 7, r.status.length, 1).setValues(r.status);
    sh.getRange(1, 7).setValue('Status (filled in automatically) - last check ' + twFmtTs_(logStamp_()) + (r.err ? ' - ⚠ ' + r.err : ''));
  } catch (x) {}
  if (r.err) return { ok: false, error: r.err };
  const old = logReadJson_(INV_FILE);
  let fresh = pre.length > 0;
  try { const s1 = tabFind_(ss, INV_TAB); if (invHeadFix_(s1)) fresh = true; } catch (x) {}
  if (old && old.v === r.cfg.v) { if (fresh) { try { invTabRefresh_(); } catch (x) {} } return { ok: true, changed: false, v: r.cfg.v, notes: pre }; }
  r.cfg.at = logStamp_(); logWriteJson_(INV_FILE, r.cfg);
  try { invTabRefresh_(); } catch (x) {}
  return { ok: true, changed: true, v: r.cfg.v, items: r.cfg.items.length, notes: pre };
}
function invSetupTab_(ss, out) {
  let sh = tabFind_(ss, INV_TAB);
  if (!sh) sh = ss.insertSheet(tabName_(INV_TAB));
  if (sh.getLastRow() >= 2) { out.push('INVENTORY tab already has rows, left as is.'); return; }
  const rows = [], seen = {}; let o = 0;
  INV_SEED.split('\n').forEach(function (l) { const i = l.indexOf('|'), c = l.slice(0, i).trim(), n = l.slice(i + 1).trim(); if (!c || seen[c]) return; seen[c] = 1; o += 10; rows.push([c, n, n, '', o, 'Yes', '']); });
  (twInvRead_().extra || []).forEach(function (x) { if (seen[x.c]) return; seen[x.c] = 1; o += 10; rows.push([x.c, x.n, x.n, INV_CATS.indexOf(String(x.cat || '').toUpperCase()) >= 0 ? String(x.cat).toUpperCase() : '', o, 'Yes', '']); });
  sh.getRange(1, 1, 1, INV_HEAD.length).setValues([INV_HEAD]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, INV_HEAD.length).setValues(rows);
  try { sh.setFrozenRows(1); [90, 280, 320, 210, 70, 70, 420].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); }); sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@'); invCatRule_(sh); sh.getRange(2, 6, Math.max(1, sh.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build()); } catch (x) {}
  out.push('INVENTORY tab created (' + rows.length + ' items).');
}
// Reorder memory from the order history file (inv-orders.json), by Item # (a rename keeps its history).
function invStats_(loc) {
  const rows = ordRows_().filter(function (r) { return !loc || String(r[1]) === String(loc); }), tz = LOG_TZ, out = {};
  const today = dayNum_(new Date(), tz);
  computeStats_(rows.map(function (r) { return [r[0], r[1], String(r[2] || '').trim().toUpperCase(), r[3], r[4]]; }), today, function (d) { return dayNum_(new Date(d), tz); }).forEach(function (s) {
    if (!s.code) return;
    out[s.code] = { times: s.times, last: s.last * 1 === s.last ? Utilities.formatDate(new Date(s.last * 86400000 + 12 * 3600000), 'UTC', 'yyyy-MM-dd') : '', avg: s.avg === null ? null : Math.round(s.avg * 10) / 10, left: s.left, since: s.since, st: s.status.indexOf('OVERDUE') === 0 ? 'due' : s.status.indexOf('ALMOST') === 0 ? 'soon' : s.status === 'OK' ? 'ok' : 'new' };
  });
  return out;
}
function invGet_(cv, loc) {
  const cfg = logReadJson_(INV_FILE);
  return { ok: true, v: cfg ? cfg.v : '', items: cfg && cfg.v !== cv ? cfg.items : undefined, stats: invStats_(loc), today: logNow_().date };
}

/* ===================== Google Sheet tab names, colors and the ALARMS tab ===================== */
// Tabs are named after the app sections (ALL CAPS). Old tab names are renamed automatically the first time they are used.
// Left = name inside this script (never shown), right = the name of the tab in the Google Sheet.
const TAB_NAMES = { 'Stock': 'INVENTORY COUNTS', 'Orders': 'INVENTORY ORDERS', 'Dashboard': 'INVENTORY DASHBOARD', 'Purchase list': 'PURCHASE LIST', 'Reports': 'MAINTENANCE REPORTS', 'Close outs': 'CLOSED REPORT',
  'Tasks': 'TASK LIST', 'Collaborators': 'COLLAB', 'PREP LIST': 'PREP LIST', 'WEEK VIEW': 'WEEK VIEW', 'FIFO HISTORY': 'FIFO REPORT', 'INVENTORY': 'INVENTORY', 'ALARMS': 'ALARMS',
  'FIFO LABELS': 'FIFO', 'TRAINING PLAN': 'TRAINING', 'TRAINING LOG': 'TRAINING LOG', 'ZELYNA Q&A': 'ZELYNA Q&A', 'ZELYNA QUESTIONS': 'ZELYNA QUESTIONS', 'Issues': 'ISSUES', 'COLLAB REPORTS': 'COLLAB REPORTS', 'RECIPES': 'RECIPES' };
// Names the tabs had before (renamed to the new name when found).
const TAB_OLD = { 'Close outs': ['CLOSE OUT CHECKLIST'], 'Tasks': ['COLLABORATOR TASK LIST'], 'Collaborators': ['COLLABORATORS'], 'PREP LIST': ["TODAY'S TASK - PREP LIST"],
  'WEEK VIEW': ["TODAY'S TASK - WEEK VIEW"], 'FIFO HISTORY': ['FIFO LABELS HISTORY'] };
function tabName_(n) { return TAB_NAMES[n] || n; }
function tabFind_(ss, n) {
  const nn = tabName_(n), c = [nn, n].concat(TAB_OLD[n] || []);
  for (let i = 0; i < c.length; i++) {
    const sh = ss.getSheetByName(c[i]);
    if (sh) { if (typeof sh.getName === 'function' && sh.getName() !== nn) { try { sh.setName(nn); } catch (e) {} } return sh; }
  }
  return null;
}
// TYPE column (Task / Cleaning): added after the last column of the TASK LIST and PREP LIST tabs, with a dropdown, so each row can be marked.
// Nothing reads it yet. It sits at the far right on purpose: the script finds every other column by its position.
const TYPE_HEAD = 'TYPE (Task / Cleaning)', TYPE_LIST = ['TASK', 'CLEANING'];
// Dropdowns (a typed value that is not in the list is still accepted; the sheet only shows a small warning):
//   FIFO tab column B (Category) = the categories of the CATEGORY DEFAULTS section (column I)
//   PREP LIST column B (WEEK VIEW category) = the WEEK VIEW sections, empty = not on the week view; column A stays free text
// Also renames the old PREP LIST title "Product (name on the FIFO label)" to "Product name". Nothing else is changed.
function dvLists_(ss, notes) {
  try {
    const fi = tabFind_(ss, LBL_TAB);
    if (fi) {
      const rows = Math.max(1, fi.getMaxRows() - LBL_FIRST + 1);
      fi.getRange(LBL_FIRST, LC.cat, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(fi.getRange(LBL_FIRST, LC.dcat, rows, 1), true).setAllowInvalid(true).build());
    }
    const pl = tabFind_(ss, PRL_TAB);
    if (pl && pl3Is_(pl)) pl.getRange(2, PL3.wc + 1, Math.max(1, pl.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(WV_CATORD.slice(), true).setAllowInvalid(true).build());
    else if (pl && pl2Is_(pl)) {
      const h = String(pl.getRange(1, 1).getValue() || '').trim();
      if (/^Product \((name on the FIFO|full name)/i.test(h)) pl.getRange(1, 1).setValue(PL2_HEAD[0]).setFontWeight('bold');
      const n = Math.max(1, pl.getMaxRows() - 1);
      pl.getRange(2, 1, n, 1).clearDataValidations();   // column A stays free text (version M had put a list there)
      pl.getRange(2, 2, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(WV_CATORD.slice(), true).setAllowInvalid(true).build());
    }
    if (notes) notes.push('Dropdowns set: FIFO Category (column B), PREP LIST WEEK VIEW.');
  } catch (e) { if (notes) notes.push('Dropdowns NOT set: ' + e); }
}
function typeColAdd_(ss, notes) {
  [[TS_TAB, TS_HEAD.length + 1, false], [PRL_TAB, PL2_HEAD.length + 1, true]].forEach(function (t) {
    try {
      const sh = tabFind_(ss, t[0]); if (!sh) return;
      if (t[2] && !pl2Is_(sh)) return;   // PREP LIST still in the old layout: do not touch
      const c = t[1], h = String(sh.getRange(1, c).getValue() || '').trim();
      if (h === TYPE_HEAD) return;
      if (h) { if (notes) notes.push(t[0] + ': column ' + c + ' already has "' + h + '", TYPE column not added.'); return; }
      sh.getRange(1, c).setValue(TYPE_HEAD).setFontWeight('bold').setBackground('#FFF2CC').setWrap(true);
      const rows = Math.max(1, sh.getMaxRows() - 1);
      sh.getRange(2, c, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TYPE_LIST, true).setAllowInvalid(false).build());
      try { sh.setColumnWidth(c, 150); sh.getRange(2, c, rows, 1).setHorizontalAlignment('center'); } catch (e) {}
      if (notes) notes.push(t[0] + ': TYPE column added (column ' + c + ', dropdown TASK / CLEANING).');
    } catch (e) { if (notes) notes.push(t[0] + ': TYPE column NOT added: ' + e); }
  });
}
/* ===================== Split rows by TYPE: CLEANING -> TASK LIST, TASK -> PREP LIST =====================
   Run from the Birriaz menu. It saves a backup of both tabs, shows what it will move, and only moves after you press Yes.
   Rows it cannot move safely stay where they are and are listed in the message (nothing is lost). */
const SPLIT_SECTION = 'PREP GENERAL';
function splitBlankTs_(r) { return r.slice(0, 14).every(function (v) { return v === '' || v === null || v === false; }); }
function splitBlankPl_(r) { return r.slice(0, 5).every(function (v) { return v === '' || v === null; }) && r.slice(5, 12).every(function (v) { return !alOnCell_(v); }) && pl2Blank_(r[12]); }
function splitType_(v) { const t = String(v || '').trim().toUpperCase(); return t === 'TASK' || t === 'CLEANING' ? t : ''; }
// Reads both tabs and decides what moves. Changes nothing.
function splitPlan_(ss, mode) {   // mode 'prep' = only move the TASK rows of the TASK LIST to the PREP LIST (nothing leaves the PREP LIST)
  const ts = tabFind_(ss, TS_TAB), pl = tabFind_(ss, PRL_TAB);
  if (!ts || !pl || !pl2Is_(pl)) return { error: 'The TASK LIST and PREP LIST tabs (new layout) are needed first.' };
  if (ts3Is_(ts) || pl3Is_(pl)) return { error: 'Not needed any more: the TASK LIST and PREP LIST have their new layout.' };
  const names = (twConfig_().employees || []).map(function (e) { return e.name; });
  const tn = Math.max(0, ts.getLastRow() - 1), pn = Math.max(0, pl.getLastRow() - 1);
  const T = tn ? ts.getRange(2, 1, tn, 16).getValues() : [], P = pn ? pl.getRange(2, 1, pn, PL2_HEAD.length + 1).getValues() : [];
  const plan = { mode: mode || '', names: names, T: T, P: P, toPrep: [], toTask: [], mergeInto: [], dropTs: [], dropPl: [], keep: [], conflicts: [], unmarked: { ts: 0, pl: 0 }, blankTsRows: [], blankPlRows: [] };
  // names already on the PREP LIST
  const byName = {}; P.forEach(function (r, i) { if (!splitBlankPl_(r)) byName[pl2Name_(String(r[0] || '').trim().toUpperCase(), String(r[1] || '').trim().toUpperCase(), String(r[4] || '').trim().toUpperCase())] = i; });
  P.forEach(function (r, i) { if (splitBlankPl_(r)) plan.blankPlRows.push(i); });
  T.forEach(function (r, i) { if (splitBlankTs_(r)) plan.blankTsRows.push(i); });
  const ticks = function (r, a) { return r.slice(a, a + 7).map(alOnCell_); };
  // 1. TASK LIST rows marked TASK -> PREP LIST
  T.forEach(function (r, i) {
    if (splitBlankTs_(r)) return;
    const ty = splitType_(r[15]), txt = String(r[2] || '').trim(), who = String(r[1] || '').trim().toUpperCase(), how = String(r[3] || '').trim().toLowerCase(), link = String(r[13] || '').split(',').map(function (x) { return x.trim().toUpperCase(); }).filter(Boolean);
    if (!ty) { plan.unmarked.ts++; return; }
    if (ty !== 'TASK') return;
    const nm = txt.toUpperCase(), id = String(r[0] || '').trim();
    if (!txt) { plan.keep.push('TASK LIST row ' + (i + 2) + ': has no task text, left as is.'); return; }
    if (who.indexOf(',') >= 0) { plan.keep.push('TASK LIST "' + txt + '": shared by several people (rotation). A prep item has one person. Left on the TASK LIST.'); return; }
    if (link.length) {
      const found = link.filter(function (k) { return byName[k] !== undefined; });
      if (found.length !== link.length) { plan.keep.push('TASK LIST "' + txt + '": its prep item "' + link.filter(function (k) { return byName[k] === undefined; }).join(', ') + '" is not on the PREP LIST. Left on the TASK LIST.'); return; }
      // the item already lives on the PREP LIST: fill what is empty there (person, days), then remove this duplicate row
      const dys = ticks(r, 4), allDays = /^(every ?day|daily)/.test(how) ? [true, true, true, true, true, true, true] : dys;
      found.forEach(function (k) { plan.mergeInto.push({ pi: byName[k], who: who, days: allDays, txt: txt }); });
      plan.dropTs.push({ i: i, id: id, why: 'duplicate of PREP LIST item ' + link.join(', ') });
      return;
    }
    let ev = 0, an = '', days = ticks(r, 4);
    const mN = /(?:every|cada)\s*(\d+)\s*(?:days?|d[ií]as?)/.exec(how);
    if (/^(every ?day|daily|diario|todos los d)/.test(how)) days = [true, true, true, true, true, true, true];
    else if (mN) {
      if (+mN[1] === 1) days = [true, true, true, true, true, true, true];
      else if (+mN[1] >= 2 && +mN[1] <= 30) { ev = +mN[1]; an = tsDate_(r[11]) || logNow_().date; days = [false, false, false, false, false, false, false]; }
      else { plan.keep.push('TASK LIST "' + txt + '": every ' + mN[1] + ' days is more than the PREP LIST allows (2 to 30). Left on the TASK LIST.'); return; }
    }
    else if (/every\s*(2|two|other)\s*weeks?|biweekly|cada\s*2\s*semanas|every other day|alternate|as needed|one ?time|once|prep/.test(how) || !/week|semana|fixed|days/.test(how)) { plan.keep.push('TASK LIST "' + txt + '": "' + String(r[3]) + '" cannot be shown as days on the PREP LIST (use weekly days, every day, or every N days). Left on the TASK LIST.'); return; }
    if (!ev && !days.some(Boolean)) { plan.keep.push('TASK LIST "' + txt + '": no days ticked. Left on the TASK LIST.'); return; }
    if (!who || names.indexOf(who) < 0) { plan.keep.push('TASK LIST "' + txt + '": the collaborator "' + who + '" is not on the COLLAB tab. Left on the TASK LIST.'); return; }
    if (byName[nm] !== undefined) { plan.keep.push('TASK LIST "' + txt + '": the PREP LIST already has an item with this name. Left on the TASK LIST.'); return; }
    plan.toPrep.push({ i: i, id: id, row: [nm.slice(0, 50), '', SPLIT_SECTION, who, '', days[0], days[1], days[2], days[3], days[4], days[5], days[6], ev || '', ev ? an : '', false, pcfgActive_(r[12]) ? 'Yes' : 'No', '', 'TASK'], txt: txt });
    byName[nm] = -1;
  });
  // 2. PREP LIST rows marked CLEANING -> TASK LIST (skipped in 'prep' mode)
  if (mode !== 'prep') P.forEach(function (r, i) {
    if (splitBlankPl_(r)) return;
    const ty = splitType_(r[PL2_HEAD.length]);
    if (!ty) { plan.unmarked.pl++; return; }
    if (ty !== 'CLEANING') return;
    const nm = pl2Name_(String(r[0] || '').trim().toUpperCase(), String(r[1] || '').trim().toUpperCase(), String(r[4] || '').trim().toUpperCase()), who = String(r[3] || '').trim().toUpperCase();
    if (!nm) { plan.keep.push('PREP LIST row ' + (i + 2) + ': empty name, left as is.'); return; }
    if (!who || names.indexOf(who) < 0) { plan.keep.push('PREP LIST "' + nm + '" (CLEANING): "Who does it" is ' + (who ? '"' + who + '", which is not a person on the COLLAB tab' : 'empty') + '. Put a person there, then run the split again. Left on the PREP LIST.'); return; }
    const days = ticks(r, 5), ev = pl2Blank_(r[12]) ? 0 : Math.round(+r[12]);
    let how, tk = days, start = '';
    if (ev >= 2 && ev <= 30) { how = 'Every ' + ev + ' days'; tk = [false, false, false, false, false, false, false]; start = tsDate_(r[13]) || ''; }
    else if (days.every(Boolean)) how = 'Every day';
    else if (days.some(Boolean)) how = 'Weekly';
    else how = 'As needed';
    plan.toTask.push({ i: i, row: ['', who, nm, how].concat(tk).concat([start, pcfgActive_(r[15]) ? 'Yes' : 'No', '', '', 'CLEANING']), nm: nm });
  });
  return plan;
}
function splitSummary_(plan) {
  const only = plan.mode === 'prep';
  const L = ['MOVE to the PREP LIST (marked TASK): ' + plan.toPrep.length + (plan.toPrep.length ? '  (' + plan.toPrep.map(function (x) { return x.txt; }).slice(0, 12).join(', ') + (plan.toPrep.length > 12 ? '...' : '') + ')' : ''),
    'REMOVE from the TASK LIST (already on the PREP LIST): ' + plan.dropTs.length,
    'Not marked yet (stay where they are): ' + (only ? plan.unmarked.ts : plan.unmarked.ts + plan.unmarked.pl)];
  if (!only) L.splice(2, 0, 'MOVE to the TASK LIST (marked CLEANING): ' + plan.toTask.length + (plan.toTask.length ? '  (' + plan.toTask.map(function (x) { return x.nm; }).slice(0, 12).join(', ') + (plan.toTask.length > 12 ? '...' : '') + ')' : ''));
  if (plan.keep.length) { L.push('', 'STAY because they cannot be moved safely (' + plan.keep.length + '):'); plan.keep.slice(0, 15).forEach(function (x) { L.push('- ' + x); }); if (plan.keep.length > 15) L.push('- ... and ' + (plan.keep.length - 15) + ' more (see the backup file).'); }
  return L.join('\n');
}
// Does the move. Saves a backup first; if the backup fails nothing is changed.
function splitRun_(ss, plan) {
  const ts = tabFind_(ss, TS_TAB), pl = tabFind_(ss, PRL_TAB), res = { prep: 0, task: 0, merged: 0, removed: 0 };
  try { logWriteJson_('type-split-backup-' + logNow_().date + '.json', { at: logStamp_(), taskList: plan.T, prepList: plan.P, kept: plan.keep }); } catch (x) { return { ok: false, error: 'Backup NOT saved, nothing was changed: ' + x }; }
  // a) people / days filled into PREP LIST items that already exist
  plan.mergeInto.forEach(function (m) {
    if (m.pi < 0 || !plan.P[m.pi]) return;
    const r = plan.P[m.pi]; let ch = false;
    if (!String(r[3] || '').trim() && m.who) { r[3] = m.who; ch = true; }
    if (!r.slice(5, 12).some(alOnCell_) && pl2Blank_(r[12]) && m.days.some(Boolean)) { m.days.forEach(function (d, k) { r[5 + k] = d; }); ch = true; }
    if (ch) { pl.getRange(m.pi + 2, 4, 1, 1).setValues([[r[3]]]); pl.getRange(m.pi + 2, 6, 1, 7).setValues([r.slice(5, 12)]); }
    else if (m.who && String(r[3] || '').trim().toUpperCase() !== m.who) plan.conflicts.push('"' + m.txt + '": the PREP LIST says ' + String(r[3]).toUpperCase() + ', the removed TASK LIST row said ' + m.who + '. The PREP LIST was kept.');
  });
  // b) write new rows into empty rows (or after the last row)
  const put = function (sh, rows, blanks, nCols, last) {
    let at = 0, next = last + 2;
    rows.forEach(function (row) {
      let r;
      if (at < blanks.length) r = blanks[at++] + 2; else { r = next++; }
      if (r > sh.getMaxRows()) { try { sh.insertRowsAfter(sh.getMaxRows(), r - sh.getMaxRows()); } catch (x) {} }
      sh.getRange(r, 1, 1, nCols).setValues([row]);
    });
  };
  put(pl, plan.toPrep.map(function (x) { return x.row; }), plan.blankPlRows, PL2_HEAD.length + 1, plan.P.length);
  put(ts, plan.toTask.map(function (x) { return x.row; }), plan.blankTsRows, 16, plan.T.length);
  res.prep = plan.toPrep.length; res.task = plan.toTask.length; res.merged = plan.mergeInto.length;
  // c) remove the moved rows from where they were (from the bottom up)
  const rm = function (sh, idx) { idx.sort(function (a, b) { return b - a; }).forEach(function (i) { sh.deleteRow(i + 2); }); };
  const tsGone = plan.toPrep.map(function (x) { return x.i; }).concat(plan.dropTs.map(function (x) { return x.i; }));
  rm(ts, tsGone); rm(pl, plan.toTask.map(function (x) { return x.i; }));
  res.removed = tsGone.length + plan.toTask.length;
  // the tablets must turn off the TASK LIST tasks that left this tab
  const ids = plan.toPrep.map(function (x) { return x.id; }).concat(plan.dropTs.map(function (x) { return x.id; })).filter(Boolean);
  try { PropertiesService.getScriptProperties().setProperty('TS_GONE_OK', ids.join(',')); } catch (x) {}
  try { tsFormat_(ts, ts.getLastRow()); } catch (x) {}
  try { pl2Format_(pl, Math.max(1, pl.getLastRow() - 1), plan.names); } catch (x) {}
  try { const dv = SpreadsheetApp.newDataValidation().requireValueInList(TYPE_LIST, true).setAllowInvalid(false).build(); ts.getRange(2, 16, Math.max(1, ts.getMaxRows() - 1), 1).setDataValidation(dv); pl.getRange(2, PL2_HEAD.length + 1, Math.max(1, pl.getMaxRows() - 1), 1).setDataValidation(dv); } catch (x) {}
  res.ok = true; res.conflicts = plan.conflicts;
  return res;
}
// Birriaz menu: Split rows by TYPE
function splitTasksToPrep() { splitUi_('prep'); }
function splitByType() { splitUi_(''); }
function splitUi_(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (x) { ui.alert('Busy', 'The sheet is updating right now. Try again in a minute.', ui.ButtonSet.OK); return; }
  let res = null, plan;
  try {
    plan = splitPlan_(ss, mode);
    if (plan.error) { ui.alert('Split rows by TYPE', plan.error, ui.ButtonSet.OK); return; }
    if (!plan.toPrep.length && !plan.toTask.length && !plan.dropTs.length) { ui.alert('Split rows by TYPE', 'Nothing to move.\n\n' + splitSummary_(plan), ui.ButtonSet.OK); return; }
    const a = ui.alert('Split rows by TYPE', splitSummary_(plan) + '\n\nA backup of both tabs is saved first. Move these rows now?', ui.ButtonSet.YES_NO);
    if (a !== ui.Button.YES) return;
    res = splitRun_(ss, plan);
  } finally { try { lock.releaseLock(); } catch (x) {} }
  if (!res || !res.ok) { ui.alert('Split rows by TYPE', (res && res.error) || 'Nothing was changed.', ui.ButtonSet.OK); return; }
  let sync = ''; try { const r = twSheetSync_(); sync = r && r.ok ? 'The tablets were updated.' : 'The tablets will update within 30 minutes.'; } catch (x) { sync = 'The tablets will update within 30 minutes.'; }
  ui.alert('Split rows by TYPE', 'Done. ' + res.prep + ' row(s) moved to the PREP LIST, ' + res.task + ' to the TASK LIST, ' + res.removed + ' removed from where they were.\n' + sync + (res.conflicts.length ? '\n\nCheck these:\n- ' + res.conflicts.join('\n- ') : ''), ui.ButtonSet.OK);
}
// Look of every tab. RED always means "do not touch":
//   red tab color   = the whole tab is filled in by the system (reports, logs, view only)
//   light red cells with a red border = columns the system fills in (status, ID, totals)
//   red line under the titles = do not rename the column titles
// Every other tab color is NOT red. [tab, tab color, row colors, columns filled in by the system or 'all', title rows]
const LOCK_RED = '#F8D7DA', LINE_RED = '#CC0000', TAB_RED = '#CC0000';
const TAB_LOOK = [
  ['INVENTORY', '#0B5394', 'LIGHT_GREY', [7, 8, 9, 10, 11, 12, 13, 14]], ['PURCHASE LIST', '#F1C232', 'LIGHT_GREY', [2, 3, 4, 7]],
  ['TASK LIST', '#6AA84F', 'LIGHT_GREY', [12, 13, 15]], ['COLLAB', '#38761D', 'LIGHT_GREY', [5]], ['PREP LIST', '#45818E', 'LIGHT_GREY', [13]],
  ['ALARMS', '#BF9000', 'LIGHT_GREY', [1, 15]], ['ZELYNA Q&A', '#F4B400', 'LIGHT_GREY', [4, 5]], ['ISSUES', TAB_RED, 'LIGHT_GREY', 'all'], ['RECIPES', '#A61C00', 'LIGHT_GREY', [7]],
  ['TRAINING', '#8E7CC3', 'LIGHT_GREY', [4, 5, 6]], ['FIFO', '#000000', null, [5], 2],
  ['WEEK VIEW', TAB_RED, null, 'all'], ['FIFO REPORT', TAB_RED, 'LIGHT_GREY', 'all'], ['CLOSED REPORT', TAB_RED, 'LIGHT_GREY', 'all'], ['MAINTENANCE REPORTS', TAB_RED, 'LIGHT_GREY', 'all'],
  ['TRAINING LOG', TAB_RED, 'LIGHT_GREY', 'all'], ['ZELYNA QUESTIONS', TAB_RED, 'LIGHT_GREY', 'all'], ['COLLAB REPORTS', TAB_RED, 'LIGHT_GREY', 'all'],
  ['INVENTORY COUNTS', TAB_RED, 'LIGHT_GREY', 'all'], ['INVENTORY ORDERS', TAB_RED, 'LIGHT_GREY', 'all'], ['INVENTORY DASHBOARD', TAB_RED, null, 'all']];
function tabLook_(sh) {
  const L = TAB_LOOK.filter(function (x) { return x[0] === sh.getName(); })[0];
  if (!L) return;
  if (tcIsSched_(sh)) { try { sh.setTabColor(L[1]); } catch (e) {} return; }   // COLLAB weekly schedule keeps its own blue look
  const hr = L[4] || 1;
  try { sh.setTabColor(L[1]); } catch (e) {}
  const cols = Math.max(1, sh.getLastColumn()), rows = Math.max(hr + 1, sh.getMaxRows());
  try {
    sh.getBandings().forEach(function (b) { b.remove(); });
    if (L[2]) sh.getRange(1, 1, rows, cols).applyRowBanding(SpreadsheetApp.BandingTheme[L[2]] || SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  } catch (e) {}
  try { sh.getRange(1, 1, 1, cols).setFontWeight('bold'); } catch (e) {}
  // red line under the column titles
  try { sh.getRange(hr, 1, 1, cols).setBorder(null, null, true, null, null, null, LINE_RED, SpreadsheetApp.BorderStyle.SOLID_THICK); } catch (e) {}
  if (L[3] === 'all') return;
  L[3].forEach(function (c) {
    if (c > cols) return;
    try {
      const r = sh.getRange(hr + 1, c, rows - hr, 1);
      r.setBackground(LOCK_RED);
      r.setBorder(null, true, null, true, null, null, LINE_RED, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      sh.getRange(1, c, hr, 1).setFontColor(LINE_RED);
    } catch (e) {}
  });
}
// Run from the Birriaz menu (or setupSheets): renames old tabs and colors every tab.
function formatSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), out = [];
  Object.keys(TAB_NAMES).forEach(function (k) { const sh = tabFind_(ss, k); if (sh) { tabLook_(sh); out.push(sh.getName()); } });
  try { if (layout3_(ss, out)) { const k = Object.keys(TAB_NAMES); k.forEach(function (x) { const sh = tabFind_(ss, x); if (sh) tabLook_(sh); }); } } catch (e) { out.push('New layout NOT made: ' + e); }
  try { CacheService.getScriptCache().remove('who3'); l3WhoDv_(ss); } catch (e) {}
  try { dvLists_(ss, out); } catch (e) {}
  try { if (!tabFind_(ss, RC_TAB)) { rcSetup_(ss); out.push('RECIPES tab created with the 14 recipes of the app.'); } } catch (e) {}
  try { const al = tabFind_(ss, AL_TAB); if (al && alIsNew_(al) && !alHas3_(al)) { alSync_(ss); out.push("ALARMS: Add to Today's task + Who does it columns added."); } } catch (e) {}
  try { hideStatus_(ss, out); } catch (e) {}
  try { rcBand_(ss, out); } catch (e) { out.push('RECIPES colors NOT applied: ' + e); }
  try { const cs = tabFind_(ss, TC_TAB); if (cs && cs.getLastRow() > 0) tcRendCol_(cs); } catch (e) {}
  // any other tab (made by hand): ALL CAPS too (formulas that point to it follow the new name by themselves)
  const known = Object.keys(TAB_NAMES).map(tabName_);
  ss.getSheets().forEach(function (sh) { const n = sh.getName(), up = n.toUpperCase(); if (known.indexOf(n) < 0 && up !== n) { const ex = ss.getSheetByName(up); if (ex && ex.getSheetId() !== sh.getSheetId()) return; try { sh.setName(up); out.push(n + ' -> ' + up); } catch (e) {} } });
  try { invTabRefresh_(); } catch (x) {}   // re-colors the Reorder status column (O) after the light red
  Logger.log('Formatted: ' + out.join(', '));
  try { ss.toast('Tabs renamed and colored. Red = do not touch.', 'Birriaz', 6); } catch (e) {}
  return out;
}

// New layout: A ID, B Time, C Alarm name, D Opens page, E..K = MON..SUN checkboxes, L Active, M Status.
// One time: the 2 new columns are inserted after Active (every row stays exactly as typed; the 3 dining checks = Yes).
function alAdd3_(sh) {
  sh.insertColumnsAfter(12, 2);
  sh.getRange(1, 13, 1, 2).setValues([["Add to Today's task", 'Who does it']]).setFontWeight('bold');
  const n = sh.getLastRow() - 1; if (n > 0) { const ids = sh.getRange(2, 1, n, 3).getValues(); sh.getRange(2, 13, n, 1).setValues(ids.map(function (r) { return [String(r[0] || r[1] || r[2] || '').trim() === '' ? '' : AL_TK_DEF.indexOf(String(r[0]).trim()) >= 0 ? 'Yes' : 'No']; })); }
  try { const m = Math.max(1, sh.getMaxRows() - 1); sh.getRange(2, 13, m, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build()); const nm = l3Names_(); if (nm.length) sh.getRange(2, 14, m, 1).setDataValidation(l3Dv_(nm)); sh.setColumnWidth(13, 120); sh.setColumnWidth(14, 140); } catch (x) {}
  try { tabLook_(sh); hideStatus_(sh.getParent(), null); } catch (x) {}
}
function alSyncGrid_(sh) {
  if (!alHas3_(sh)) { try { alAdd3_(sh); } catch (x) { Logger.log('ALARMS columns not added: ' + x); } }
  const has3 = alHas3_(sh), stc = has3 ? 15 : 13, names = l3Names_();
  const n = sh.getLastRow() - 1, rows = n > 0 ? sh.getRange(2, 1, n, has3 ? 14 : 12).getValues() : [], st = [], seen = {}, list = [], off = {};
  rows.forEach(function (r, i) {
    const blank = function (v) { return String(v === null || v === false ? '' : v).trim() === ''; };
    if (blank(r[0]) && blank(r[1]) && blank(r[2])) { st.push(['']); return; }
    let id = String(r[0] || '').trim();
    const t = alParseT_(r[1]), nm = String(r[2] || '').trim().toUpperCase().slice(0, 40), d = [], er = [];
    r.slice(4, 11).forEach(function (v, k) { if (alOnCell_(v)) d.push(WV_COLS[k]); });
    if (!t) er.push('Time not understood (use 1:15 PM)'); if (!nm) er.push('Alarm name is empty'); if (!d.length) er.push('Tick at least one day');
    if (id && seen[id]) er.push('This ID is used twice. Clear the ID on the new row');
    if (er.length) { st.push(['⚠ ' + er.join('; ') + ' (not used)']); return; }
    if (!id) { id = 'k' + Date.now().toString(36) + i; try { sh.getRange(i + 2, 1).setValue(id); } catch (x) {} }
    seen[id] = 1; d.sort();
    const tk = has3 ? pcfgActive_(r[12]) && String(r[12] === null ? '' : r[12]).trim() !== '' : AL_TK_DEF.indexOf(id) >= 0;
    let who = has3 ? String(r[13] || '').trim().toUpperCase().slice(0, 40) : '', wn = '';
    if (who && names.length && names.indexOf(who) < 0) { wn = ' · ⚠ ' + who + ' is not on the COLLAB tab (left empty)'; who = ''; }
    list.push({ id: id, t: t, n: nm, d: d, go: alGoOf_(r[3]), tk: tk, who: who });
    if (!pcfgActive_(r[11])) off[id] = 1;
    st.push([(off[id] ? 'Off (does not ring)' : '✓ OK') + (tk && !off[id] ? " · adds to Today's task" + (who ? ' for ' + who : '') : '') + wn]);
  });
  try { if (st.length) sh.getRange(2, stc, st.length, 1).setValues(st); sh.getRange(1, stc).setValue('Status (filled in automatically) - last check ' + twFmtTs_(logStamp_())); } catch (x) {}
  if (!list.length) return { ok: false, error: 'The ALARMS tab has no alarms, so nothing was changed on the tablets.' };
  return alSave_(list, off);
}
const AL_TAB = 'ALARMS', AL_HEAD = ['ID (do not change)', 'Time', 'Alarm name', 'Days', 'Opens page', 'Active', 'Status (filled in automatically)'];
const AL_GO = { '': '', 'close': 'Close out checklist', 'temps': 'Food temperature log', 'prep': "Today's task" };
const AL_SEED = [['10:22','SET REGISTER AND FLAGS',[0,6]],['10:25','SET FOOD WARMERS TO #4',[0,6]],['10:55','SET FOOD WARMERS TO #4',[1,2,3,4,5]],['10:55','SET REGISTER AND FLAGS',[1,2,3,4,5]],['11:00','LOG TEMPS',null,'temps'],['12:00','ADD CONSOME'],['13:00','LOG TEMPS',null,'temps'],['14:00','ADD CONSOME'],['15:00','LOG TEMPS',null,'temps'],['15:50','CHANGE WATER BUCKETS'],['16:00','ADD CONSOME'],['17:00','LOG TEMPS',null,'temps'],['18:00','ADD CONSOME'],['19:00','LOG TEMPS',null,'temps'],['20:13','CLOSE OUT CHECKLIST',[1,2,3],'close'],['21:48','CLOSE OUT CHECKLIST',[4,5,6,0],'close'],['13:15','CHECK AND SWEEP DINING AREA (CASHIER)',null,'prep'],['16:15','CHECK AND SWEEP DINING AREA (CASHIER)',null,'prep'],['18:15','CHECK AND SWEEP DINING AREA (CASHIER)',null,'prep']];
const AL_DN = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
function alSeed_() { return AL_SEED.map(function (a, i) { return { id: 'a' + i, t: a[0], n: a[1], d: a[2] || [0, 1, 2, 3, 4, 5, 6], go: a[3] || '', tk: AL_TK_DEF.indexOf('a' + i) >= 0, who: '' }; }); }
function alList_(doc) { return (doc.list && doc.list.length) ? doc.list : alSeed_().concat((doc.custom || []).map(function (c) { return { id: c.id, t: c.t, n: c.n, d: c.d, go: '' }; })); }
function alT12_(t) { const h = +t.slice(0, 2); return ((h % 12) || 12) + ':' + t.slice(3) + ' ' + (h < 12 ? 'AM' : 'PM'); }
function alDaysTxt_(d) { const k = d.slice().sort().join(''); return k === '0123456' ? 'Every day' : k === '12345' ? 'Weekdays' : k === '06' ? 'Weekends' : d.slice().sort(function (a, b) { return (a || 7) - (b || 7); }).map(function (x) { return AL_DN[x]; }).join(', '); }
function alParseT_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') { return ('0' + v.getHours()).slice(-2) + ':' + ('0' + v.getMinutes()).slice(-2); }
  const m = /^\s*(\d{1,2})(?::(\d{2}))?\s*(am|pm|a\.m\.|p\.m\.)?\s*$/i.exec(String(v || ''));
  if (!m) return null; let h = +m[1]; const mi = +(m[2] || 0), ap = (m[3] || '').toLowerCase();
  if (ap.charAt(0) === 'p' && h < 12) h += 12; if (ap.charAt(0) === 'a' && h === 12) h = 0;
  if (h > 23 || mi > 59) return null; return ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2);
}
function alParseDays_(v) {
  const p = tsPlain_(v).trim();
  if (!p || /^(EVERY ?DAY|DAILY|DIARIO|TODOS)/.test(p)) return [0, 1, 2, 3, 4, 5, 6];
  if (/^WEEKDAYS|ENTRE SEMANA/.test(p)) return [1, 2, 3, 4, 5];
  if (/^WEEKENDS|FIN DE SEMANA|FINES/.test(p)) return [0, 6];
  return wvDays_(p);
}
function alGoOf_(v) { const s = String(v || '').toLowerCase(); return /close|cierre/.test(s) ? 'close' : /temp/.test(s) ? 'temps' : /today|task|prep|tarea/.test(s) ? 'prep' : ''; }
function alSetupTab_(ss, out) {
  let sh = tabFind_(ss, AL_TAB);
  if (!sh) sh = ss.insertSheet(tabName_(AL_TAB));
  if (sh.getLastRow() >= 2) { out.push('ALARMS tab already has rows, left as is.'); return; }
  const doc = logReadJson_('alarms.json') || {};
  alSheetWrite_(sh, alList_(doc), doc.off || {}, '');
  try { sh.setFrozenRows(1); [140, 110, 330, 200, 200, 70, 380].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); }); } catch (x) {}
  out.push('ALARMS tab created.');
}
function alSheetWrite_(sh, list, off, note) {
  const rows = list.map(function (a) { return [a.id, alT12_(a.t), a.n, AL_GO[a.go || ''] || ''].concat(WV_COLS.map(function (d) { return a.d.indexOf(d) >= 0; })).concat([off[a.id] ? 'No' : 'Yes', alTk_(a) ? 'Yes' : 'No', a.who || '', note || '']); });
  try { sh.getBandings().forEach(function (b) { b.remove(); }); } catch (x) {}
  try { sh.getRange(1, 1, Math.max(1, sh.getMaxRows()), Math.max(15, sh.getMaxColumns ? sh.getMaxColumns() : 15)).clearDataValidations(); } catch (x) {}
  sh.clear();
  sh.getRange(1, 1, 1, AL_HEAD2.length).setValues([AL_HEAD2]).setFontWeight('bold');
  const extra = 15, n = rows.length + extra;
  try { sh.getRange(2, 2, n, 1).setNumberFormat('@'); } catch (x) {}
  if (rows.length) sh.getRange(2, 1, rows.length, AL_HEAD2.length).setValues(rows);
  try {
    sh.getRange(2, 5, n, 7).insertCheckboxes();
    if (rows.length) sh.getRange(2, 5, rows.length, 7).setValues(rows.map(function (r) { return r.slice(4, 11); }));
    sh.getRange(2, 4, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['', 'Close out checklist', 'Food temperature log', "Today's task"], true).build());
    sh.getRange(2, 12, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
    sh.getRange(2, 13, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
    try { const nm = l3Names_(); if (nm.length) sh.getRange(2, 14, n, 1).setDataValidation(l3Dv_(nm)); } catch (x) {}
    sh.setFrozenRows(1); [140, 100, 330, 190, 55, 55, 55, 55, 55, 55, 55, 70, 120, 140, 380].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange(1, 5, n + 1, 7).setHorizontalAlignment('center');
  } catch (x) {}
}
// ALARMS tab -> alarms.json (the tablets pull it every minute). Never wipes: an empty tab changes nothing.
function alSync_(ss) {
  const sh = tabFind_(ss, AL_TAB);
  if (!sh) return { ok: false, error: 'no ALARMS tab yet (run setupSheets once)' };
  if (alIsNew_(sh)) return alSyncGrid_(sh);
  const n = sh.getLastRow() - 1, rows = n > 0 ? sh.getRange(2, 1, n, 6).getValues() : [], st = [], seen = {}, list = [], off = {};
  rows.forEach(function (r, i) {
    if (r.every(function (v) { return String(v === null ? '' : v).trim() === ''; })) { st.push(['']); return; }
    let id = String(r[0] || '').trim();
    const t = alParseT_(r[1]), nm = String(r[2] || '').trim().toUpperCase().slice(0, 40), d = alParseDays_(r[3]), er = [];
    if (!t) er.push('Time not understood (use 1:15 PM)'); if (!nm) er.push('Alarm name is empty'); if (!d.length) er.push('Days not understood (use Every day, Weekdays, Weekends or Mon, Thu)');
    if (id && seen[id]) er.push('This ID is used twice. Clear the ID on the new row');
    if (er.length) { st.push(['⚠ ' + er.join('; ') + ' (not used)']); return; }
    if (!id) { id = 'k' + Date.now().toString(36) + i; try { sh.getRange(i + 2, 1).setValue(id); } catch (x) {} }
    seen[id] = 1;
    list.push({ id: id, t: t, n: nm, d: d, go: alGoOf_(r[4]) });
    if (!pcfgActive_(r[5])) off[id] = 1;
    st.push([off[id] ? 'Off (does not ring)' : '✓ OK']);
  });
  try { if (st.length) sh.getRange(2, 7, st.length, 1).setValues(st); sh.getRange(1, 7).setValue('Status (filled in automatically) - last check ' + twFmtTs_(logStamp_())); } catch (x) {}
  if (!list.length) return { ok: false, error: 'The ALARMS tab has no alarms, so nothing was changed on the tablets.' };
  try { alSheetWrite_(sh, list, off, ''); tabLook_(sh); } catch (x) {}   // old layout -> one row per alarm with day checkboxes
  return alSave_(list, off);
}
function alSave_(list, off) {
  const cur = logReadJson_('alarms.json') || {};
  if (JSON.stringify(cur.list || null) === JSON.stringify(list) && JSON.stringify(cur.off || {}) === JSON.stringify(off)) return { ok: true, changed: false };
  logWriteJson_('alarms.json', { t: Date.now(), list: list, off: off, from: 'sheet' });
  return { ok: true, changed: true, alarms: list.length };
}

/* ---- WEEK VIEW tab as a weekly grid (same layout as the app): categories down the side, MONDAY ... SUNDAY across ---- */
// Type an item under its day, in its category block. Empty rows in each block are for new items. Move an item = cut and paste it to another day.
const WV_COLS = [1, 2, 3, 4, 5, 6, 0], WV_HDR = ['CATEGORY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY', '', 'Status (filled in automatically)'];
const WV_CATORD = ['PICAR (DICED MEAT)', 'VEGGIES PREPARATION', 'COSINAR (COOK ITEMS)', 'PREP GENERAL', 'SAUCE PREP'];
const WV_CATCOL = { 'PICAR (DICED MEAT)': '#F6D5CC', 'VEGGIES PREPARATION': '#CDE9EE', 'SAUCE PREP': '#E3D8F0', 'COSINAR (COOK ITEMS)': '#FFE0B2', 'PREP GENERAL': '#F5D3E3' };
const WV_EVERY = 'EVERY FEW DAYS', WV_WHO = 'WHO DOES IT (label on Today\'s task)';
function wvIsGrid_(sh) { return String(sh.getRange(1, 1).getValue() || '').trim().toUpperCase() === 'CATEGORY'; }
const WV_VIEW_NOTE = "VIEW ONLY - made automatically from the PREP LIST tab. To change a day, tick it there.";
// view = true: the WEEK VIEW is only a picture of the PREP LIST (no spare rows, no WHO section). Returns false when nothing had to change.
function wvWriteGrid_(sh, weekly, every, lbl, plItems, view) {
  const cats = WV_CATORD.slice(); weekly.forEach(function (d) { d.forEach(function (w) { if (cats.indexOf(w[0]) < 0) cats.push(w[0]); }); });
  const rows = [view ? WV_HDR.slice(0, 9).concat([WV_VIEW_NOTE]) : WV_HDR], blocks = [];
  cats.forEach(function (c) {
    const col = WV_COLS.map(function (d) { return weekly[d].filter(function (w) { return w[0] === c; }).map(function (w) { return w[1] + (view && lbl[w[1]] ? ' (' + lbl[w[1]] + ')' : ''); }); });
    const mx = Math.max.apply(null, col.map(function (x) { return x.length; }));
    const n = view ? Math.max(1, mx) : Math.max(3, mx + 2), start = rows.length + 1;
    if (!col.some(function (x) { return x.length; }) && (view || WV_CATORD.indexOf(c) > 3)) return;   // empty sections are left out (SAUCE PREP only when it has items)
    for (let r = 0; r < n; r++) rows.push([r === 0 ? c : ''].concat(col.map(function (x) { return x[r] || ''; })).concat(['', '']));
    blocks.push([c, start, n]); rows.push(['', '', '', '', '', '', '', '', '', '']);
  });
  const evStart = rows.length + 1;
  rows.push(view ? [WV_EVERY, 'Item', 'Section', 'Every (days)', 'Start date', '', '', '', '', ''] : [WV_EVERY, 'Item', 'Category', 'Every (days)', 'Start date', 'Active', '', '', '', '']);
  every.forEach(function (x) { rows.push(['', x.n, x.c, x.e, x.an, view ? '' : 'Yes', '', '', '', '']); });
  let whoStart = 0;
  if (!view) {
    for (let i = 0; i < 2; i++) rows.push(['', '', '', '', '', '', '', '', '', '']);
    rows.push(['', '', '', '', '', '', '', '', '', '']);
    whoStart = rows.length + 1; const pl = {}; (plItems || []).forEach(function (n) { pl[n] = 1; });
    rows.push([WV_WHO, 'Item', 'Who does it', '', '', '', '', '', '', '']);
    const seen = {}; weekly.forEach(function (d) { d.forEach(function (w) { seen[w[1]] = 1; }); }); every.forEach(function (x) { seen[x.n] = 1; });
    Object.keys(seen).filter(function (n) { return !pl[n]; }).sort().forEach(function (n) { rows.push(['', n, lbl[n] || '', '', '', '', '', '', '', '']); });
  }
  rows.push(['', '', '', '', '', '', '', '', '', '']);
  if (view) {
    // nothing to do when the tab already shows exactly this (dates compared as dates)
    try {
      const norm = function (v) { return Object.prototype.toString.call(v) === '[object Date]' || /^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? 'D' : String(v === null || v === undefined ? '' : v); };
      let last = rows.length; while (last > 1 && rows[last - 1].every(function (v) { return v === ''; })) last--;
      if (sh.getLastRow() === last && String(sh.getRange(1, 10).getValue() || '') === WV_VIEW_NOTE) {
        const cur = sh.getRange(1, 1, last, 9).getValues();
        if (cur.every(function (r, i) { return r.every(function (v, j) { return norm(v) === norm(rows[i][j]); }); })) return false;
      }
    } catch (x) {}
  }
  try { sh.getBandings().forEach(function (b) { b.remove(); }); } catch (x) {}
  try { sh.getRange(1, 1, Math.max(1, sh.getMaxRows()), Math.max(10, sh.getMaxColumns ? sh.getMaxColumns() : 10)).breakApart(); } catch (x) {}
  try { sh.getRange(1, 1, Math.max(1, sh.getMaxRows()), Math.max(10, sh.getMaxColumns ? sh.getMaxColumns() : 10)).clearDataValidations(); } catch (x) {}
  sh.clear();
  sh.getRange(1, 1, rows.length, 10).setValues(rows);
  try {
    sh.setFrozenRows(1); sh.setColumnWidth(1, 150); for (let c = 2; c <= 8; c++) sh.setColumnWidth(c, 190); sh.setColumnWidth(9, 20); sh.setColumnWidth(10, 360);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#D9D9D9');
    blocks.forEach(function (b) {
      const col = WV_CATCOL[b[0]] || '#EEEEEE', rg = sh.getRange(b[1], 1, b[2], 8);
      rg.setBackground(col).setBorder(true, true, true, true, true, true);
      sh.getRange(b[1], 1, b[2], 1).merge().setFontWeight('bold').setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
    });
    sh.getRange(evStart, 1, 1, view ? 5 : 6).setFontWeight('bold').setBackground('#D9D9D9'); if (whoStart) sh.getRange(whoStart, 1, 1, 3).setFontWeight('bold').setBackground('#D9D9D9');
    sh.getRange(evStart + 1, 5, Math.max(1, every.length + 2), 1).setNumberFormat('m/d/yyyy');
    if (view) {
      sh.getRange(1, 10).setFontWeight('bold').setWrap(true).setBackground(LOCK_RED);
      sh.getRange(2, 1, Math.max(1, rows.length - 1), 8).setWrap(true);
      try { sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); }); sh.protect().setDescription('WEEK VIEW is made from the PREP LIST tab').setWarningOnly(true); } catch (x) {}
    }
  } catch (x) {}
  return true;
}
// Reads the grid. Returns {weekly, every, lbl, st} where st = status text per row (row 2 on).
function wvParseGrid_(sh, today) {
  const n = sh.getLastRow() - 1, rows = n > 0 ? sh.getRange(2, 1, n, 8).getValues() : [], weekly = [[], [], [], [], [], [], []], every = [], lbl = {}, st = [];
  let mode = 'grid', cat = '';
  rows.forEach(function (r) {
    const a = String(r[0] === null ? '' : r[0]).trim(), A = a.toUpperCase(), cells = r.slice(1).map(function (v) { return String(v === null || v === false ? '' : v).trim(); });
    if (A === WV_EVERY) { mode = 'every'; st.push(['']); return; }
    if (A.indexOf('WHO DOES IT') === 0) { mode = 'who'; st.push(['']); return; }
    if (mode === 'grid') {
      if (a) cat = A.slice(0, 60);
      const items = cells.map(function (v) { return v.toUpperCase().slice(0, 60); });
      if (!items.some(Boolean)) { st.push(['']); return; }
      if (!cat) { st.push(['⚠ This row has no category. Put the category in column A of its block (not used)']); return; }
      items.forEach(function (it, i) { if (it) weekly[WV_COLS[i]].push([cat, it]); });
      st.push(['✓ OK']); return;
    }
    if (mode === 'every') {
      const it = cells[0].toUpperCase().slice(0, 60), c = cells[1].toUpperCase().slice(0, 60);
      if (/^ITEM$/i.test(cells[0]) || (!it && !c)) { st.push(['']); return; }
      const k = Math.round(+cells[2]), er = []; let an = tsDate_(r[4]);
      if (!it) er.push('Item is empty'); if (!c) er.push('Category is empty'); if (!(k >= 2 && k <= 30)) er.push('Every (days) must be 2 to 30'); if (an === null) er.push('Start date not understood (use 11/3/2026)');
      if (er.length) { st.push(['⚠ ' + er.join('; ') + ' (not used)']); return; }
      if (!pcfgActive_(r[5])) { st.push(['Inactive (hidden on the tablets)']); return; }
      every.push({ c: c, n: it, e: k, an: an || today }); st.push(['✓ OK']); return;
    }
    const it = cells[0].toUpperCase().slice(0, 60), lb = cells[1].toUpperCase().slice(0, 30);
    if (/^ITEM$/i.test(cells[0]) || !it) { st.push(['']); return; }
    if (lb) lbl[it] = lb; st.push(['✓ OK']);
  });
  return { weekly: weekly, every: every, lbl: lbl, st: st };
}

/* ---- ALARMS tab, horizontal: one row per alarm, a checkbox under each day ---- */
const AL_HEAD2 = ['ID (do not change)', 'Time', 'Alarm name', 'Opens page', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN', 'Active', "Add to Today's task", 'Who does it', 'Status (filled in automatically)'];
// Alarms that put a task on Today's task when they ring. Alarms saved before this column existed: the 3 dining-area checks (1, 4 and 6 PM).
const AL_TK_DEF = ['a16', 'a17', 'a18'];
function alTk_(a) { return a && a.tk !== undefined ? !!a.tk : AL_TK_DEF.indexOf(String(a && a.id)) >= 0; }
function alHas3_(sh) { return /^Add to Today/i.test(String(sh.getRange(1, 13).getValue() || '').trim()); }
function alIsNew_(sh) { return String(sh.getRange(1, 5).getValue() || '').trim().toUpperCase() === 'MON'; }
function alOnCell_(v) { return v === true || /^(true|x|✓|✔|yes|si|sí|1)$/i.test(String(v || '').trim()); }

/*PREPSHARED-START*/
const PREPS=[['PICAR (DICED MEAT)',['ASADA','POLLO','ALPASTOR']],
['VEGGIES PREPARATION',['CILANTRO','LIMONES','CEBOLLA BLANCA','CEBOLLA MORADA','TOMATO','PICO','FAJITAS','REPOLLO','LECHUGA','CEBOLLA PARA ALPASTOR','GUACAMOLE']],
['SAUCE PREP',['PINK SAUCE','BAJA SAUCE','GUAC SAUCE','CREMA']],
['COSINAR (COOK ITEMS)',['ARROS BLANCO','ARROS MEXICANO','NACHOS','FRIJOLES','ACEITE ROJO','GRILL CHICKEN','GRILL ASADA','GRILL ALPASTOR','SALSA ROJA','SALSA VERDE','ROSTISAR CHILES','POTATO','TAMALES']],
['PREP GENERAL',['BATIDAS DE ARROZ','MARINAR ASADA','MARINAR AL PASTOR','GUARDAR ACEITE','GUARDAR FRIJOLES','PAPEL EN CUADRITOS']]];
const PITEMS=[];PREPS.forEach(g=>g[1].forEach(n=>PITEMS.push({id:'p_'+n,n:n,g:g[0]})));
const WEEKLY=[[['VEGGIES PREPARATION','WASH CILANTRO TODO'],['VEGGIES PREPARATION','LIMONES TODO'],['PREP GENERAL','MARINAR ASADA'],['PREP GENERAL','TAMALES PREP'],['COSINAR (COOK ITEMS)','COOK CHICKEN AM 11'],['COSINAR (COOK ITEMS)','NACHOS'],['COSINAR (COOK ITEMS)','MEXICAN RICE']],
[['VEGGIES PREPARATION','GUACAMOLE'],['COSINAR (COOK ITEMS)','PASTOR ASAR AM 11'],['COSINAR (COOK ITEMS)','BIRRIA'],['PICAR (DICED MEAT)','PICAR POLLO PM']],
[['VEGGIES PREPARATION','PICAR CILANTRO'],['PREP GENERAL','ROAST CHILES'],['COSINAR (COOK ITEMS)','ASADA ASAR AM 11'],['COSINAR (COOK ITEMS)','BIRRIA'],['PICAR (DICED MEAT)','PICAR AL PASTOR']],
[['VEGGIES PREPARATION','WASH CILANTRO TODO'],['VEGGIES PREPARATION','LIMONES TODO'],['VEGGIES PREPARATION','PICO 1 CONT'],['VEGGIES PREPARATION','TOMATE 1 CONT'],['VEGGIES PREPARATION','CEBOLLA ALPASTOR 1 CONT'],['COSINAR (COOK ITEMS)','BIRRIA'],['PICAR (DICED MEAT)','PICAR ASADA']],
[['VEGGIES PREPARATION','WHITE ONION TODO'],['VEGGIES PREPARATION','PURPLE ONION TODO'],['VEGGIES PREPARATION','PICAR CILANTRO TODO'],['COSINAR (COOK ITEMS)','SALSAS'],['COSINAR (COOK ITEMS)','TAMALES COOK']],
[['VEGGIES PREPARATION','PICAR CILANTRO'],['VEGGIES PREPARATION','FAJITAS 2-3 CONT'],['COSINAR (COOK ITEMS)','BEANS'],['COSINAR (COOK ITEMS)','WHITE RICE']],
[['VEGGIES PREPARATION','PICO 1 CONT'],['VEGGIES PREPARATION','TOMATE 1 CONT'],['VEGGIES PREPARATION','GUAC SAUCE'],['VEGGIES PREPARATION','GUACAMOLE'],['VEGGIES PREPARATION','PINK SAUCE'],['PREP GENERAL','MARINAR ALPASTOR'],['COSINAR (COOK ITEMS)','OIL'],['COSINAR (COOK ITEMS)','POTATO']]];
WEEKLY.unshift(WEEKLY.pop());WEEKLY[1].push(['VEGGIES PREPARATION','GREEN ONION']);WEEKLY[0].push(['PREP GENERAL','BATIDAS DE ARROZ']);
// Items that come back every N days (not tied to a weekday). Before 'an' (start date) the old weekdays 'old' (0=Sun) are used.
const PEVERY=[{c:'VEGGIES PREPARATION',n:'LECHUGA 2 CONT',e:3,an:'2026-10-31',old:[1,6]}];
function pDn_(ds){const p=ds.split('-');return Math.round(Date.UTC(+p[0],+p[1]-1,+p[2])/864e5)}
function pEv_(ds){return PEVERY.filter(x=>ds<x.an?x.old.indexOf(pDow_(ds))>=0:(pDn_(ds)-pDn_(x.an))%x.e===0)}
function pDow_(ds){const p=ds.split('-');return new Date(Date.UTC(+p[0],+p[1]-1,+p[2])).getUTCDay()}
// Start a new day for one location: unfinished items carry over as URGENT, finished ones drop, today's weekly items are added.
function pRoll_(st,pw,ds){
if(!st.n)st.n={};if(!st.x)st.x={};if(!st.u)st.u={};if(!st.c)st.c=[];if(!st.o)st.o={};if(!st.e)st.e={};
if(st.d){Object.keys(st.n).forEach(id=>{if(st.x[id]||id.indexOf('dn_')===0){delete st.n[id];delete st.x[id];delete st.u[id];delete st.e[id]}else st.u[id]=1});st.c=st.c.filter(c=>st.n[c.id])}
st.d=ds;st.o={};st.b={};const dw=pDow_(ds);
WEEKLY[dw].forEach((w,i)=>{const m=PITEMS.find(p=>p.n===w[1]);if(m){st.n[m.id]=1;return}
if(st.c.some(c=>c.n===w[1]))return;const id='w'+dw+'_'+i;st.c.push({id:id,n:w[1],g:'WEEKLY: '+w[0]});st.n[id]=1});
pEv_(ds).forEach(x=>{const m=PITEMS.find(p=>p.n===x.n);if(m){st.n[m.id]=1;return}if(st.c.some(c=>c.n===x.n))return;const id='e_'+x.n;st.c.push({id:id,n:x.n,g:'WEEKLY: '+x.c});st.n[id]=1});
const pm=pw[ds]||{};PITEMS.concat(st.c).forEach(i=>{if(st.n[i.id]&&pm[i.n]){st.x[i.id]=1;delete st.u[i.id]}});
const lim=logAddDays_(ds,-35);Object.keys(pw).forEach(k=>{if(k<lim)delete pw[k]})}
// Apply ONE change. k = n:/x:/u:/c:<id>, o:<n>|o:s (close-out), w:<date>|<NAME> (week view mark). v = 0/1 (c: {n,g} or 0)
function pSet_(st,pw,k,v){const p=k.charAt(0),id=k.slice(2);if(!st.b)st.b={};if(!st.e)st.e={};
if(p==='n'||p==='x'||p==='u'){if(v)st[p][id]=1;else delete st[p][id]}
else if(p==='b'){if(v)st.b[id]=String(v).slice(0,10);else delete st.b[id]}
else if(p==='e'){if(v)st.e[id]=String(v).slice(0,30);else delete st.e[id]}
else if(p==='o'){if(v)st.o[id]=1;else delete st.o[id]}
else if(p==='c'){st.c=st.c.filter(c=>c.id!==id);if(v)st.c.push({id:id,n:String(v.n||'').slice(0,60),g:String(v.g||'').slice(0,80)})}
else if(p==='w'){const i=id.indexOf('|'),d=id.slice(0,i),nm=id.slice(i+1);if(v){(pw[d]=pw[d]||{})[nm]=1}else if(pw[d]){delete pw[d][nm];if(!Object.keys(pw[d]).length)delete pw[d]}}}
// The Prep list and Week view can come from the Google Sheet (PREP LIST and WEEK VIEW tabs).
// cfg = {v, items:[{n,g,lbl,sl}], weekly:[ [[category,name,label],...] for Sun..Sat ], every:[{c,n,e,an,lbl}]}. Without a cfg the built-in lists above are used.
function pApplyCfg_(cfg){if(!cfg||!Array.isArray(cfg.items)||!cfg.items.length||!Array.isArray(cfg.weekly))return false;
const by={},ord=[];cfg.items.forEach(i=>{if(!by[i.g]){by[i.g]=[];ord.push(i.g)}by[i.g].push(i.n)});
PREPS.length=0;ord.forEach(g=>PREPS.push([g,by[g]]));
PITEMS.length=0;PREPS.forEach(g=>g[1].forEach(n=>PITEMS.push({id:'p_'+n,n:n,g:g[0]})));
for(let d=0;d<7;d++)WEEKLY[d]=(cfg.weekly[d]||[]).map(w=>[w[0],w[1]]);
PEVERY.length=0;(cfg.every||[]).forEach(x=>PEVERY.push({c:x.c,n:x.n,e:x.e,an:x.an,old:[]}));return true}
/*PREPSHARED-END*/

/* ===================== ZELYNA: the AI assistant (text + voice), powered by Claude through THIS secure backend ===================== */
// The tablets never see the API key. The page sends a question here; this code asks Claude and lets it LOOK THINGS UP with
// read-only tools. The location comes from the request and is forced on every lookup, files are only read from the allowed
// SOP folder(s), and employee information needs the manager code. Settings (Project Settings > Script properties):
//   ZELYNA_AI_KEY        the Anthropic API key (the old NOVI_AI_KEY name keeps working)
//   ZELYNA_MODEL         optional, default claude-sonnet-5-5
//   ZELYNA_LOCATIONS     optional, comma list of the locations of THIS restaurant (default: the ones already in the inventory data)
//   ZELYNA_DOC_FOLDERS   optional, comma list of Drive folder ids Zelyna may read (default: the folder named BIRRIAZ SOPs)
//   ZELYNA_MAX_DAY       optional, questions per location per day (default 300)
const ZEL_MODEL_DEFAULT = 'claude-sonnet-5-5', ZEL_SOP_FOLDER = 'BIRRIAZ SOPs';
const ZEL_KEYS_OPEN = ['INVENTORY', 'Purchase list', 'ALARMS', 'PREP LIST', 'Reports', 'Close outs', 'FIFO HISTORY', 'FIFO LABELS'], ZEL_KEYS_MGR = ['Tasks', 'Collaborators'];
const ZEL_TABS_OPEN = ZEL_KEYS_OPEN.map(tabName_), ZEL_TABS_MGR = ZEL_KEYS_MGR.map(tabName_);
function zelProp_(k) { try { return PropertiesService.getScriptProperties().getProperty(k) || ''; } catch (e) { return ''; } }
function zelKey_() { return zelProp_('ZELYNA_AI_KEY') || zelProp_('NOVI_AI_KEY'); }
function zelModel_() { return zelProp_('ZELYNA_MODEL') || ZEL_MODEL_DEFAULT; }
function zelLocs_() {
  const p = zelProp_('ZELYNA_LOCATIONS');
  if (p) return p.split(',').map(function (x) { return x.trim(); }).filter(String);
  const s = {};
  try { Object.keys(cntRead_().by).forEach(function (l) { s[l] = 1; }); } catch (e) {}
  try { ordRows_().forEach(function (r) { if (r[1]) s[String(r[1])] = 1; }); } catch (e) {}
  return Object.keys(s);
}
function zelLocOk_(loc) {
  if (!loc || !/^[\w .&'’-]{1,50}$/.test(loc)) return false;
  const l = zelLocs_(); return !l.length || l.indexOf(loc) >= 0;
}
// Questions per location per day (protects the AI bill when many tablets use it).
function zelCap_(loc) {
  const P = PropertiesService.getScriptProperties(), today = logNow_().date, max = Number(zelProp_('ZELYNA_MAX_DAY')) || 300;
  let u = {}; try { u = JSON.parse(P.getProperty('ZEL_USE') || '{}'); } catch (e) {}
  if (u.d !== today) u = { d: today, n: {} };
  if ((u.n[loc] || 0) >= max) return false;
  u.n[loc] = (u.n[loc] || 0) + 1; P.setProperty('ZEL_USE', JSON.stringify(u));
  return true;
}
function zelCell_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v) ? Utilities.formatDate(v, LOG_TZ, 'yyyy-MM-dd h:mm a') : (v === null || v === undefined ? '' : (typeof v === 'string' ? v.slice(0, 300) : v)); }

/* ===================== ZELYNA KNOWLEDGE: reference Sheets + a fast search index =====================
   Run from the Birriaz menu: "Build Zelyna knowledge (Sheets + search)". It reads every Google Doc, PDF and Word file in the
   authorized folder (BIRRIAZ SOPs and its subfolders) and writes, word for word, one Google Sheet per category in
   BIRRIAZ SOPs > ZELYNA KNOWLEDGE > <category>. One tab per document, one row per instruction (section, step, exact text,
   temperatures / times, warnings). The original documents are never changed, moved or deleted.
   The same rows feed a search index (a file in the data folder + a short memory cache), so Zelyna answers from one quick
   lookup instead of opening documents. Every 30 minutes only the documents that changed are converted again. */
const ZKB_ROOT = 'ZELYNA KNOWLEDGE', ZKB_FILE = 'zelyna-knowledge-index.json', ZKB_STATE = 'zelyna-knowledge-sheets.json';
const ZKB_CATS = [
  { k: 'daily_operations', name: 'DAILY OPERATIONS', folder: '1 DAILY OPERATIONS', re: /dress|uniform|attendance|tardi|break|shift|reminder|deposit|cashier|register|disciplin|policy|opening|closing|schedule|horario|uniforme|asistencia|turno/i },
  { k: 'sops', name: 'STANDARD OPERATING PROCEDURES', folder: '2 SOPs', re: null },
  { k: 'food_safety', name: 'FOOD SAFETY', folder: '3 FOOD SAFETY', re: /food safety|hand ?wash|glove|cooking food|cooling|thaw|temperatur|reheat|allergen|seguridad alimentaria|guantes|enfriamiento/i },
  { k: 'cleaning', name: 'CLEANING PROCEDURES', folder: '4 CLEANING', re: /clean|dish|bathroom|restroom|hood|sanitiz|trash|dirty|limpi|trastes|baño|basura/i },
  { k: 'training', name: 'TRAINING MATERIALS', folder: '5 TRAINING', re: /training|course|food handler|manipulador|quiz|exam|orientation|onboard|capacitaci|curso/i },
  { k: 'approved_qa', name: 'APPROVED Q&A', folder: '6 APPROVED Q&A', re: null },
  { k: 'menu', name: 'MENU / RECIPES', folder: '7 MENU', re: /\b(menu|men[uú]|recipes?|recetas?|build|platillos?)\b/i }];
const ZKB_SCOPE = { daily_operations: ['daily_operations'], sops: ['sops'], food_safety: ['food_safety', 'training'], cleaning: ['cleaning'], training: ['training', 'food_safety'], menu: ['menu'] };
const ZKB_HEAD = ['Section', 'Sub-section', 'Step', 'Instruction (exact text)', 'Temperatures / times', 'Important', 'Note'];
function zkbCatOf_(title) {
  const t = String(title || '');
  if (/template/i.test(t)) return '';
  const order = ['menu', 'training', 'cleaning', 'food_safety', 'daily_operations'];
  for (let i = 0; i < order.length; i++) { const c = ZKB_CATS.filter(function (x) { return x.k === order[i]; })[0]; if (c.re.test(t)) return c.k; }
  return 'sops';
}
function zkbCat_(k) { return ZKB_CATS.filter(function (x) { return x.k === k; })[0]; }
// temperatures, times and warning words found in one line
function zkbKeys_(s) {
  const out = [], t = String(s || '');
  (t.match(/\d{2,3}\s*°\s*[FC]?(\s*(–|-|to|a)\s*\d{2,3}\s*°\s*[FC]?)?/gi) || []).forEach(function (x) { out.push(x.replace(/\s+/g, ' ')); });
  (t.match(/\b\d+(\.\d+)?\s*(hours?|hrs?|minutes?|mins?|seconds?|secs?|days?|weeks?|horas?|minutos?|segundos?|d[ií]as?|semanas?)\b/gi) || []).forEach(function (x) { out.push(x.replace(/\s+/g, ' ')); });
  return out.filter(function (x, i) { return out.indexOf(x) === i; }).join(', ');
}
function zkbImportant_(s) { return /\b(never|do not|don't|must|always|only|warning|danger|discard|throw (it )?out|immediately|nunca|no (se )?debe|siempre|solo|peligro|desechar|inmediatamente)\b|°/i.test(String(s || '')) ? 'YES' : ''; }
function zkbEsTab_(title, text) {
  if (/spanish|español|espanol/i.test(title)) return true;
  const s = ' ' + String(text || '').slice(0, 3000).toLowerCase() + ' ';
  const es = (s.match(/\s(el|la|los|las|de|del|para|con|una|que|en|por|antes|después|lavar|manos)\s/g) || []).length, en = (s.match(/\s(the|and|of|to|with|before|after|for|your|hands|wash)\s/g) || []).length;
  return es > en * 1.3;
}
// ---- one Google Doc -> rows, keeping every word (headings, steps and sub-steps, tables row by row)
function zkbDocRows_(id, wantPh) {
  const doc = DocumentApp.openById(id), tabs = [], ph = [];
  const walk = function (list) { list.forEach(function (tb) { try { tabs.push({ title: tb.getTitle(), body: tb.asDocumentTab().getBody() }); } catch (e) {} try { walk(tb.getChildTabs ? tb.getChildTabs() : []); } catch (e) {} }); };
  try { walk(doc.getTabs()); } catch (e) {}
  if (!tabs.length) tabs.push({ title: '', body: doc.getBody() });
  const rows = [], ET = DocumentApp.ElementType, PH = DocumentApp.ParagraphHeading;
  tabs.forEach(function (tb, ti) {
    let sec = '', sub = '', n = [0, 0, 0, 0, 0, 0], tabText = '';
    try { tabText = tb.body.getText(); } catch (e) {}
    const es = zkbEsTab_(tb.title, tabText), lang = es ? 'ES' : 'EN';
    if (tabs.length > 1) rows.push({ lang: lang, sec: (tb.title || ('Tab ' + (ti + 1))).toUpperCase(), sub: '', step: '', x: '', head: 1 });
    const num = tb.body.getNumChildren();
    for (let i = 0; i < num; i++) {
      const el = tb.body.getChild(i), ty = el.getType();
      let note = '';
      if (ty === ET.PARAGRAPH || ty === ET.LIST_ITEM) {
        const p = ty === ET.PARAGRAPH ? el.asParagraph() : el.asListItem();
        try { const k = p.getNumChildren(); for (let j = 0; j < k; j++) { const c = p.getChild(j).getType(); if (wantPh && c === ET.INLINE_IMAGE && ph.length < 12) { try { ph.push({ b: p.getChild(j).asInlineImage().getBlob(), sec: sec, sub: sub }); note = '📷 Photo saved (Zelyna can show it)'; continue; } catch (e) {} } if (c === ET.INLINE_IMAGE || c === ET.INLINE_DRAWING) note = '⚠ Picture in the original (not converted). Open the document to see it.'; if (c === ET.EQUATION) note = '⚠ Formula in the original (check the document).'; } } catch (e) {}
        try { if (p.getPositionedImages && p.getPositionedImages().length) note = '⚠ Picture in the original (not converted). Open the document to see it.'; } catch (e) {}
        const x = String(p.getText() || '').replace(/ /g, ' ').replace(/\s+$/g, '').trim();
        if (!x || /^[-–—_=*\s]{3,}$/.test(x)) { if (note) rows.push({ lang: lang, sec: sec, sub: sub, step: '', x: '(picture)', note: note }); continue; }
        if (ty === ET.LIST_ITEM) {
          const lv = Math.min(5, Math.max(0, el.asListItem().getNestingLevel()));
          n[lv]++; for (let q = lv + 1; q < 6; q++) n[q] = 0;
          rows.push({ lang: lang, sec: sec, sub: sub, step: n.slice(0, lv + 1).map(function (v) { return v || 1; }).join('.'), x: x, note: note, lv: lv });
          continue;
        }
        const h = p.getHeading();
        let isHead = h !== PH.NORMAL && h !== undefined && h !== null, big = h === PH.TITLE || h === PH.HEADING1 || h === PH.HEADING2;
        if (!isHead) { try { const tx = p.editAsText(); isHead = x.length <= 80 && !/[.!?]$/.test(x) && tx.isBold(0) === true && tx.isBold(x.length - 1) === true && x.indexOf(':') !== x.length - 2; } catch (e) {} }
        if (isHead) {
          if (big || !sec) { sec = x; sub = ''; } else sub = x;
          n = [0, 0, 0, 0, 0, 0];
          rows.push({ lang: lang, sec: sec, sub: sub, step: '', x: x, head: 1, note: note });
        } else rows.push({ lang: lang, sec: sec, sub: sub, step: '', x: x, note: note });
      } else if (ty === ET.TABLE) {
        const t = el.asTable();
        for (let r = 0; r < t.getNumRows(); r++) {
          const row = t.getRow(r), cells = [];
          for (let c = 0; c < row.getNumCells(); c++) cells.push(String(row.getCell(c).getText() || '').replace(/\s+/g, ' ').trim());
          if (cells.join('')) rows.push({ lang: lang, sec: sec, sub: sub, step: '', x: cells.join(' | '), note: r === 0 ? 'Table in the original (written here row by row, cells split by | ).' : '' });
        }
      } else if (ty === ET.INLINE_DRAWING || ty === ET.INLINE_IMAGE) rows.push({ lang: lang, sec: sec, sub: sub, step: '', x: '(picture)', note: '⚠ Picture in the original (not converted). Open the document to see it.' });
    }
  });
  rows.ph = ph;
  return rows;
}
// ---- PDF / Word: the text Drive reads from it (pictures and layout are not kept, so every row says so once)
function zkbTextRows_(text, kind) {
  const rows = []; let sec = '', n = 0;
  String(text || '').split(/\r?\n/).forEach(function (l, i) {
    const x = l.replace(/ /g, ' ').trim(); if (!x) return;
    const head = x.length <= 70 && (/^[A-ZÁÉÍÓÚÑ0-9 #&()\-–:,'/]+$/.test(x) && /[A-Z]{3}/.test(x) || /:$/.test(x));
    if (head) { sec = x.replace(/:$/, ''); n = 0; rows.push({ lang: zkbEsTab_('', x) ? 'ES' : 'EN', sec: sec, sub: '', step: '', x: x, head: 1 }); return; }
    const m = /^(\d+[.)]|[-•●▪*])\s+(.*)$/.exec(x);
    if (m) { n++; rows.push({ lang: '', sec: sec, sub: '', step: String(n), x: m[2] }); } else rows.push({ lang: '', sec: sec, sub: '', step: '', x: x });
  });
  if (rows.length) rows[0].note = '⚠ Made from a ' + kind + ': pictures, columns and layout are not kept. Check the original for those.';
  return rows;
}
function zkbFileRows_(f, wantPh) {
  const kind = zelDocKind_(f.getMimeType());
  if (kind === 'Google Doc') return zkbDocRows_(f.getId(), wantPh);
  if (kind === 'PDF' || kind === 'Word') return zkbTextRows_(zelConvText_(f), kind);
  if (kind === 'Text') return zkbTextRows_(f.getBlob().getDataAsString('UTF-8'), 'text file');
  return [];
}
// ---- index: rows grouped by section into short searchable pieces
function zkbChunks_(rows) {
  const out = []; let cur = null;
  const flush = function () { if (cur && cur.lines.length) { const x = cur.lines.join('\n'); for (let i = 0; i < x.length; i += 1500) out.push({ s: cur.s, l: cur.l, x: x.slice(i, i + 1600) }); } cur = null; };
  rows.forEach(function (r) {
    const s = [r.sec, r.sub].filter(Boolean).join(' › ');
    if (!cur || cur.s !== s || (r.head && !r.sub && r.sec !== cur.sec)) { flush(); cur = { s: s, sec: r.sec, l: r.lang || '', lines: [] }; }
    if (r.x && !(r.head && (r.x === r.sec || r.x === r.sub))) cur.lines.push((r.step ? r.step + '. ' : '') + r.x);
  });
  flush();
  return out.filter(function (c) { return c.x.trim(); });
}
function zkbIdxRead_() {
  try { const c = CacheService.getScriptCache(), m = c.get('zkb:n'); if (m) { const keys = []; for (let i = 0; i < +m; i++) keys.push('zkb:' + i); const all = c.getAll(keys); if (Object.keys(all).length === +m) return JSON.parse(keys.map(function (k) { return all[k]; }).join('')); } } catch (e) {}
  const d = logReadJson_(ZKB_FILE); if (d && d.docs) zkbCache_(d);
  return d && d.docs ? d : null;
}
function zkbCache_(d) {
  try { const s = JSON.stringify(d), c = CacheService.getScriptCache(), o = {}; let n = 0; for (let i = 0; i < s.length; i += 90000) o['zkb:' + (n++)] = s.slice(i, i + 90000); if (n <= 90) { c.putAll(o, 21600); c.put('zkb:n', String(n), 21600); } } catch (e) {}
}
function zkbIdxWrite_(d) { d.at = logStamp_(); logWriteJson_(ZKB_FILE, d); try { CacheService.getScriptCache().remove('zkb:n'); } catch (e) {} zkbCache_(d); }
function zkbNorm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9°ñ ]+/g, ' '); }
const ZKB_STOP = { the: 1, a: 1, an: 1, and: 1, or: 1, of: 1, to: 1, in: 1, on: 1, for: 1, is: 1, are: 1, do: 1, does: 1, we: 1, i: 1, you: 1, how: 1, what: 1, when: 1, where: 1, which: 1, can: 1, it: 1, my: 1, our: 1, with: 1, be: 1, at: 1, el: 1, la: 1, los: 1, las: 1, de: 1, del: 1, y: 1, o: 1, que: 1, en: 1, un: 1, una: 1, para: 1, con: 1, por: 1, es: 1, se: 1, como: 1, cuando: 1, donde: 1, cual: 1, cuanto: 1, cuantos: 1, mi: 1, nuestro: 1, hay: 1, me: 1, lo: 1, al: 1 };
function zkbTok_(s) { return zkbNorm_(s).split(/\s+/).filter(function (w) { return w && !ZKB_STOP[w] && (w.length > 1 || /\d/.test(w)); }).map(function (w) { return w.length > 4 ? w.replace(/(es|s)$/, '') : w; }); }
// Search: best sections of the knowledge sheets, only in the category asked (falls back to all when nothing is found there)
function zkbSearch_(query, cat) {
  const d = zkbIdxRead_();
  if (!d || !d.docs || !Object.keys(d.docs).length) return { results: [], note: 'The knowledge index is not built yet. Use search_documents.' };
  const q = zkbTok_(query); if (!q.length) return { results: [] };
  const es = zkbEsTab_('', query);
  const all = [], df = d.df || {};
  Object.keys(d.docs).forEach(function (id) { const D = d.docs[id]; (D.c || []).forEach(function (c) { all.push({ id: id, D: D, c: c, toks: c.w !== undefined ? String(c.w).split(' ') : zkbTok_(D.t + ' ' + c.s + ' ' + c.x), head: c.h !== undefined ? String(c.h).split(' ') : zkbTok_(D.t + ' ' + c.s) }); }); });
  const N = d.N || all.length || 1;
  const score = function (it) {
    let sc = 0; q.forEach(function (w) { const tf = it.toks.filter(function (x) { return x === w || (w.length > 4 && x.indexOf(w) === 0); }).length; if (!tf) return; const idf = Math.log(1 + N / (df[w] || 1)); sc += idf * (1 + Math.log(tf)) + (it.head.indexOf(w) >= 0 ? idf * 2 : 0); });
    if (!sc) return 0;
    const hit = q.filter(function (w) { return it.toks.some(function (x) { return x === w || (w.length > 4 && x.indexOf(w) === 0); }); }).length;
    sc *= 0.5 + hit / q.length;
    if (it.c.l && (it.c.l === 'ES') === es) sc *= 1.15;
    return sc * (it.D.p || 1);
  };
  const scope = cat && ZKB_SCOPE[cat] ? ZKB_SCOPE[cat] : null;
  let pool = scope ? all.filter(function (it) { return scope.indexOf(it.D.cat) >= 0; }) : all, widened = false;
  let best = pool.map(function (it) { return { it: it, s: score(it) }; }).filter(function (x) { return x.s > 0; });
  if (!best.length && scope) { widened = true; best = all.map(function (it) { return { it: it, s: score(it) }; }).filter(function (x) { return x.s > 0; }); }
  best.sort(function (a, b) { return b.s - a.s; });
  const res = [], per = {};
  best.forEach(function (b) { if (res.length >= 3 || (per[b.it.id] || 0) >= 2) return; per[b.it.id] = (per[b.it.id] || 0) + 1; const pr = { document: b.it.D.t, category: (zkbCat_(b.it.D.cat) || {}).name || '', type: b.it.D.k, section: b.it.c.s, text: b.it.c.x, document_id: b.it.id }, ph = zkbPhotoFor_(b.it.D, b.it.c.s); if (ph.length) pr.photo_ids = ph; res.push(pr); });
  return { results: res, searched: widened ? 'all categories (nothing in the one asked)' : (scope ? scope.map(function (k) { return zkbCat_(k).name; }).join(', ') : 'all categories'), built: d.at };
}
// ---- sheets
function zkbFolder_(parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function zkbTabName_(t, used) { let n = String(t || 'Document').replace(/[\[\]\*\?\/\\:]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 90) || 'Document', k = n, i = 2; while (used[k.toUpperCase()]) k = n.slice(0, 85) + ' (' + (i++) + ')'; used[k.toUpperCase()] = 1; return k; }
function zkbBook_(st, root, cat) {
  const C = zkbCat_(cat); let ss = null;
  if (st.books[cat]) { try { ss = SpreadsheetApp.openById(st.books[cat]); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('ZELYNA · ' + C.name + ' (made automatically)');
    try { DriveApp.getFileById(ss.getId()).moveTo(zkbFolder_(zkbFolder_(root, ZKB_ROOT), C.folder)); } catch (e) {}
    st.books[cat] = ss.getId();
  }
  return ss;
}
function zkbWriteTab_(ss, name, title, link, head, rows, firstNote) {
  let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  const w = head.length, data = [[firstNote || ('Made automatically from "' + title + '". Do not edit here: change the original document; this tab follows it within 30 minutes.')].concat(new Array(w - 1).fill('')), [link || ''].concat(new Array(w - 1).fill('')), head].concat(rows.map(function (r) { const o = r.slice(0, w); while (o.length < w) o.push(''); return o; }));
  sh.getRange(1, 1, data.length, w).setValues(data);
  try {
    sh.getRange(1, 1, 2, w).setFontColor('#B71C1C').setFontWeight('bold'); sh.getRange(3, 1, 1, w).setFontWeight('bold').setBackground('#F8D7DA').setBorder(null, null, true, null, null, null, '#CC0000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.setFrozenRows(3); sh.setTabColor('#CC0000');
    if (rows.length) sh.getRange(4, 1, rows.length, w).setWrap(true).setVerticalAlignment('top');
    [180, 170, 60, 520, 170, 90, 260].slice(0, w).forEach(function (px, i) { sh.setColumnWidth(i + 1, px); });
  } catch (e) {}
  return sh;
}
function zkbDocRowsOut_(rows) {
  return rows.map(function (r) { return [r.sec || '', r.sub || '', r.step || '', r.x || '', zkbKeys_(r.x), r.head ? '' : zkbImportant_(r.x), r.note || '']; });
}
// ---- build (all) or refresh (only the documents that changed)
function zkbList_() {
  const out = [], seen = {};
  zelAllFolders_().forEach(function (fid) {
    try { const fo = DriveApp.getFolderById(fid), inMenu = /^(menu|men[uú]|recipes?|recetas?)\b/i.test(fo.getName()), it = fo.getFiles(); while (it.hasNext() && out.length < 160) { const f = it.next(), k = zelDocKind_(f.getMimeType()); if (!k || seen[f.getId()]) continue; seen[f.getId()] = 1; let cat = zkbCatOf_(f.getName()); if (cat && inMenu) cat = 'menu';   /* a file in the MENU folder always goes to the menu file */
      if (cat) out.push({ f: f, id: f.getId(), t: f.getName(), k: k, cat: cat, upd: String(f.getLastUpdated().getTime()) }); } } catch (e) {}
  });
  return out;
}
function zkbRun_(full, maxMs) {
  const t0 = Date.now(), lim = maxMs || 270000, roots = zelFolders_();
  if (!roots.length) return { ok: false, error: 'The BIRRIAZ SOPs folder was not found.' };
  const root = DriveApp.getFolderById(roots[0]);
  const idx = (!full && zkbIdxRead_()) || { v: 1, docs: {} }, st = logReadJson_(ZKB_STATE) || { books: {}, tabs: {} };
  st.books = st.books || {}; st.tabs = st.tabs || {};
  const list = zkbList_(), res = { ok: true, converted: [], unchanged: 0, removed: 0, flagged: [], failed: [], left: 0 };
  const here = {}; list.forEach(function (x) { here[x.id] = 1; });
  Object.keys(idx.docs).forEach(function (id) { if (!here[id]) { delete idx.docs[id]; res.removed++; zkbPhotosDrop_(st, id); } });
  const todo = list.filter(function (x) { const D = idx.docs[x.id]; return full || !D || D.u !== x.upd || D.cat !== x.cat; });
  res.unchanged = list.length - todo.length;
  const rowsBy = {};
  for (let i = 0; i < todo.length; i++) {
    if (Date.now() - t0 > lim) { res.left = todo.length - i; break; }
    const x = todo[i];
    try {
      const rows = zkbFileRows_(x.f, x.cat === 'menu'), out = zkbDocRowsOut_(rows), flags = rows.filter(function (r) { return /⚠/.test(r.note || ''); }).length;
      const ph = zkbPhotos_(st, root, x, rows.ph || []);
      const link = x.f.getUrl ? x.f.getUrl() : '';
      const ch = zkbChunks_(rows); ch.forEach(function (c) { c.w = zkbTok_(x.t + ' ' + c.s + ' ' + c.x).join(' '); c.h = zkbTok_(x.t + ' ' + c.s).join(' '); });
      idx.docs[x.id] = { t: x.t, k: x.k, cat: x.cat, u: x.upd, link: link, p: x.k === 'Google Doc' ? (/master|480 sop/i.test(x.t) ? 0.85 : 1) : 0.8, c: ch, n: rows.length, fl: flags, ph: ph };
      rowsBy[x.id] = { x: x, out: out, link: link, flags: flags };
      res.converted.push(x.t); if (flags) res.flagged.push(x.t + ' (' + flags + ')');
    } catch (e) { res.failed.push(x.t + ': ' + String(e).slice(0, 120)); }
  }
  // word counts for the ranking, computed once here (not on every question)
  idx.df = {}; idx.N = 0;
  Object.keys(idx.docs).forEach(function (id) { (idx.docs[id].c || []).forEach(function (c) { idx.N++; const set = {}; String(c.w || '').split(' ').forEach(function (w) { if (w) set[w] = 1; }); Object.keys(set).forEach(function (w) { idx.df[w] = (idx.df[w] || 0) + 1; }); }); });
  zkbIdxWrite_(idx);
  // sheets: one book per category, one tab per document + CONTENTS (+ FOOD TEMPERATURES, TRAINING Q&A, APPROVED Q&A)
  try {
    Object.keys(rowsBy).forEach(function (id) {
      const R = rowsBy[id], ss = zkbBook_(st, root, R.x.cat), used = {};
      Object.keys(st.tabs).forEach(function (k) { if (k !== id && st.tabs[k].cat === R.x.cat) used[String(st.tabs[k].name).toUpperCase()] = 1; });
      ['CONTENTS', 'FOOD TEMPERATURES', 'TRAINING Q&A', 'APPROVED Q&A'].forEach(function (k) { used[k] = 1; });
      const old = st.tabs[id];
      if (old && old.cat !== R.x.cat) { try { const o = SpreadsheetApp.openById(st.books[old.cat]).getSheetByName(old.name); if (o) o.getParent().deleteSheet(o); } catch (e) {} }
      const name = old && old.cat === R.x.cat ? old.name : zkbTabName_(R.x.t, used);
      zkbWriteTab_(ss, name, R.x.t, R.link, ZKB_HEAD, R.out.length ? R.out : [['', '', '', '(no text found in this document)', '', '', '⚠ Check the original']]);
      st.tabs[id] = { cat: R.x.cat, name: name };
    });
    Object.keys(st.tabs).forEach(function (id) { if (!idx.docs[id]) { const o = st.tabs[id]; try { const sh = SpreadsheetApp.openById(st.books[o.cat]).getSheetByName(o.name); if (sh) sh.getParent().deleteSheet(sh); } catch (e) {} delete st.tabs[id]; } });
    if (Object.keys(rowsBy).length || res.removed || full) zkbSummaryTabs_(st, root, idx);
  } catch (e) { res.failed.push('Sheets: ' + String(e).slice(0, 160)); }
  logWriteJson_(ZKB_STATE, st);
  res.docs = Object.keys(idx.docs).length;
  return res;
}
// Photos of a menu Doc: saved as image files in ZELYNA KNOWLEDGE / PHOTOS (the old ones of that Doc are replaced when the Doc changes).
// Each photo keeps the heading it was under, so "DINE IN" and "TO GO" photos stay with their own steps.
function zkbPhotosDrop_(st, docId) { st.ph = st.ph || {}; (st.ph[docId] || []).forEach(function (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} }); delete st.ph[docId]; }
function zkbPhotos_(st, root, x, list) {
  zkbPhotosDrop_(st, x.id);
  if (!list.length) return [];
  const fo = zkbFolder_(zkbFolder_(root, ZKB_ROOT), 'PHOTOS'), out = [];
  list.forEach(function (p, i) {
    try {
      const b = p.b, ext = /png/i.test(b.getContentType() || '') ? 'png' : 'jpg';
      const f = fo.createFile(b.setName(String(x.t).replace(/[\\/:*?"<>|]+/g, ' ').slice(0, 80) + ' - photo ' + (i + 1) + '.' + ext));
      out.push({ f: f.getId(), s: [p.sec, p.sub].filter(Boolean).join(' › ') });
    } catch (e) { Logger.log('Photo not saved (' + x.t + '): ' + e); }
  });
  st.ph = st.ph || {}; st.ph[x.id] = out.map(function (p) { return p.f; });
  return out;
}
// The photos that fit one search result: the ones under the same heading; when the photo has its own heading (PHOTO), all of the Doc's photos.
function zkbPhotoFor_(D, section) {
  const ph = D.ph || []; if (!ph.length) return [];
  const top = String(section || '').split(' › ')[0].toUpperCase();
  const same = ph.filter(function (p) { return top && String(p.s || '').split(' › ')[0].toUpperCase() === top && !/^(PHOTOS?|FOTOS?|PICTURES?)$/.test(top); });
  return (same.length ? same : ph).slice(0, 2).map(function (p) { return p.f; });
}
function zkbSummaryTabs_(st, root, idx) {
  ZKB_CATS.forEach(function (C) {
    const ids = Object.keys(idx.docs).filter(function (id) { return idx.docs[id].cat === C.k; });
    if (!ids.length && C.k !== 'approved_qa' && !st.books[C.k]) return;
    const ss = zkbBook_(st, root, C.k);
    const rows = ids.map(function (id) { const D = idx.docs[id]; return [D.t, D.k, (st.tabs[id] || {}).name || '', D.n || 0, D.fl ? D.fl + ' ⚠' : '', D.link || '']; });
    zkbWriteTab_(ss, 'CONTENTS', '', '', ['Document', 'Type', 'Tab here', 'Rows', 'Items to check in the original', 'Original document'], rows,
      C.name + ' — made automatically from the BIRRIAZ SOPs folder. Do not edit here: change the original documents. Originals are never changed.');
    try { const sh = ss.getSheetByName('CONTENTS'); ss.setActiveSheet(sh); ss.moveActiveSheet(1); const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1'); if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1); } catch (e) {}
  });
  // food temperatures from every document
  try {
    const t = [];
    Object.keys(idx.docs).forEach(function (id) { const D = idx.docs[id]; (D.c || []).forEach(function (c) { c.x.split('\n').forEach(function (l) { if (/\d{2,3}\s*°/.test(l)) t.push([D.t, c.s, l, zkbKeys_(l), D.link || '']); }); }); });
    zkbWriteTab_(zkbBook_(st, root, 'food_safety'), 'FOOD TEMPERATURES', '', '', ['Document', 'Section', 'Exact text', 'Temperatures / times', 'Original document'], t,
      'Every line with a temperature, taken word for word from the documents. Do not edit here.');
  } catch (e) {}
  // approved Q&A (copy of the ZELYNA Q&A tab) and training questions (weekly lessons)
  try {
    const qa = zelFaqRows_().filter(function (r) { return String(r[0] || '').trim() && zelOn_(r[2]); }).map(function (r) { return [String(r[0]), String(r[1] || '(owner note)'), String(r[3] || '')]; });
    zkbWriteTab_(zkbBook_(st, root, 'approved_qa'), 'APPROVED Q&A', '', '', ['Question', 'Answer', 'Category'], qa, 'Copy of the ZELYNA Q&A tab (the tab is the original; change answers there).');
    const tr = []; try { const d = tnRead_(); Object.keys(d.weeks || {}).sort().forEach(function (w) { const W = d.weeks[w]; ['en', 'es'].forEach(function (l) { const L = W && W[l]; if (L && L.q) tr.push([w, l.toUpperCase(), L.title || '', L.q, L.a || '']); }); }); } catch (e) {}
    zkbWriteTab_(zkbBook_(st, root, 'training'), 'TRAINING Q&A', '', '', ['Week of', 'Language', 'Lesson', 'Question', 'Answer'], tr, 'Questions of the weekly trainings, made automatically. Do not edit here.');
  } catch (e) {}
}
// Birriaz menu
function zkbBuildMenu() {
  const r = zkbRun_(true, 280000);
  try { zkbTrigger_(); } catch (e) {}
  const ui = SpreadsheetApp.getUi();
  const L = ['Documents converted: ' + r.converted.length + (r.left ? ' (' + r.left + ' more will finish in the next minutes automatically)' : ''), 'Search index: ' + (r.docs || 0) + ' documents'];
  if (r.flagged.length) L.push('', 'Check in the original (pictures or tables): ' + r.flagged.join(', '));
  if (r.failed.length) L.push('', 'Could not convert: ' + r.failed.join('; '));
  L.push('', 'The new Sheets are in BIRRIAZ SOPs > ' + ZKB_ROOT + '. The original documents were not changed.');
  if (r.left) { try { ScriptApp.newTrigger('zkbRefresh').timeBased().after(60000).create(); } catch (e) {} }
  ui.alert('Zelyna knowledge', r.ok ? L.join('\n') : r.error, ui.ButtonSet.OK);
}
// every 30 minutes: only what changed
function zkbRefresh() {   // never uses the lock the tablets save with; a flag keeps two refreshes from running at once
  const c = CacheService.getScriptCache(); if (c.get('zkb:run')) return; c.put('zkb:run', '1', 300);
  try { zkbRun_(false, 240000); } finally { try { c.remove('zkb:run'); } catch (e) {} }
}
function zkbTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'zkbRefresh') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('zkbRefresh').timeBased().everyMinutes(30).create();
}

const ZEL_CATS = ['Food safety', 'Prep and recipes', 'Cleaning', 'Equipment', 'Inventory and orders', 'Customer service', 'Cash and register', 'Schedule and team', 'Opening and closing', 'Other'];
const ZEL_TOOLS = [
  { name: 'inventory_lookup', description: 'Look up inventory items of THIS location by name, item number or category (leave query empty to list what is marked as out / needed right now). Returns the count marked out here and the reorder memory (last order, times ordered, days left, status).', input_schema: { type: 'object', properties: { query: { type: 'string' } } } },
  { name: 'temperature_status', description: 'Food temperature log for the current week: compliance, failed readings with their corrective actions, and checks missing or pending today.', input_schema: { type: 'object', properties: {} } },
  { name: 'task_status', description: 'Collaborator tasks for the current week: per employee scheduled / done / pending / overdue, and the open items. MANAGER MODE ONLY.', input_schema: { type: 'object', properties: { employee: { type: 'string' } } } },
  { name: 'recent_records', description: 'Recent maintenance reports, close out checklists and inventory orders sent for THIS location.', input_schema: { type: 'object', properties: { kind: { type: 'string', enum: ['maintenance', 'closeout', 'orders'] } }, required: ['kind'] } },
  { name: 'read_sheet', description: 'Read the first rows of an authorized Google Sheet tab. Tabs: ' + ZEL_TABS_OPEN.join(', ') + ' (manager mode also: ' + ZEL_TABS_MGR.join(', ') + ').', input_schema: { type: 'object', properties: { tab: { type: 'string' } }, required: ['tab'] } },
  { name: 'get_recipe', description: 'The batch recipes of the restaurant (RECIPES tab of the Google Sheet): ingredients, measurement, amount and notes, exactly as the owner wrote them. Use it for how much of each ingredient goes in a recipe (pico, guac sauce, horchata, rice...). Without a name it lists the recipe names.', input_schema: { type: 'object', properties: { name: { type: 'string' } } } },
  { name: 'search_knowledge', description: 'FIRST place to look after the approved Q&A. Searches the restaurant knowledge sheets (made word for word from the SOPs, food safety, cleaning, daily operations and training documents) and returns the best matching sections with their exact text. Choose the category of the question; leave it out only when unsure.', input_schema: { type: 'object', properties: { query: { type: 'string' }, category: { type: 'string', enum: ZKB_CATS.filter(function (c) { return c.k !== 'approved_qa'; }).map(function (c) { return c.k; }) } }, required: ['query'] } },
  { name: 'search_documents', description: 'Only when search_knowledge did not have the answer or enough detail: search the original SOPs, Google Docs, PDFs and Word files in the authorized folder and its subfolders (names and text). Returns document names, ids and types. If nothing matches, it also lists the available documents so you can pick one.', input_schema: { type: 'object', properties: { query: { type: 'string' } }, required: ['query'] } },
  { name: 'read_document', description: 'Read the full text of one authorized document (a document_id from search_knowledge or an id from search_documents) when the section found is not enough: Google Doc, PDF or Word file. Each SOP Google Doc has an English and a Spanish tab.', input_schema: { type: 'object', properties: { id: { type: 'string' } }, required: ['id'] } },
  { name: 'log_unanswered_question', description: 'Save a question you could not answer so the owner can add the answer to the ZELYNA Q&A tab. Use the question as the person asked it.', input_schema: { type: 'object', properties: { question: { type: 'string' }, category: { type: 'string', enum: ZEL_CATS } }, required: ['question'] } }
];

function zelInventory_(loc, q) {
  const cfg = logReadJson_(INV_FILE) || {}, items = cfg.items || [], cn = cntIndex_(), st = invStats_(loc), k = String(q || '').trim().toLowerCase();
  const rows = items.map(function (i) {
    const c = String(i.c).toUpperCase(), o = cn.code[c], n = o && o.locs[loc] ? o.locs[loc] : 0, s = st[c];
    return { code: i.c, name: i.n, category: i.cat || '', marked_out: n, reorder_status: s ? s.st : '', times_ordered: s ? s.times : 0, last_order: s ? s.last : '', days_left: s ? s.left : null };
  });
  const hit = k ? rows.filter(function (r) { return String(r.code).toLowerCase() === k || r.name.toLowerCase().indexOf(k) >= 0 || r.category.toLowerCase().indexOf(k) >= 0; }) : rows.filter(function (r) { return r.marked_out > 0; });
  return { location: loc, count: hit.length, items: hit.slice(0, 25) };
}
function zelTemps_() {
  const c = logConfig_(), now = logNow_(), wk = logMonday_(now.date), d = logReadWeek_(wk, false), S = logStats_(c, d.entries, wk, now);
  const today = S.days.filter(function (x) { return x.date === now.date; })[0], open = today ? today.slots.filter(function (s) { return s.status === 'missing' || s.status === 'pending'; }) : [];
  return { week: wk, compliance_percent: S.compliance, readings_done: S.done, failed: S.failed, missing: S.missing, pending: S.pending,
    open_today: open.slice(0, 20).map(function (s) { return { check: s.check, equipment: s.equip, food: s.food, status: s.status }; }),
    failed_readings: d.entries.filter(function (e) { return e.result === 'FAIL'; }).slice(-15).map(function (e) { return { date: e.date, check: e.check, equipment: e.equip, food: e.food, temp: e.temp, action: e.action, who: e.who }; }) };
}
function zelTasks_(emp) {
  const cfg = twConfig_(), now = logNow_(), wk = logMonday_(now.date), d = twReadWeek_(wk, false), S = twStats_(cfg, d.comps, twCarry_(cfg, wk), wk, now, false);
  const e = String(emp || '').trim().toUpperCase(), names = S.order.filter(function (n) { return !e || n.toUpperCase() === e; });
  return { week: wk, today: now.date, employees: names.map(function (n) { const a = S.byEmp[n]; return { name: n, scheduled: a.sched, done: a.done, pending: a.pending, overdue: a.overdue, completed_percent: a.pct }; }),
    open_items: S.items.filter(function (x) { return (x.st === 'pending' || x.st === 'overdue') && (!e || String(x.emp).toUpperCase() === e); }).slice(0, 30).map(function (x) { const t = twFind_(cfg, x.tid); return { employee: x.emp, task: t ? t.en : x.tid, due: x.due, status: x.st, days_late: x.late }; }) };
}
function zelRecords_(loc, kind) {
  const mine = function (rows, ci) { return rows.filter(function (r) { return String(r[ci]) === String(loc); }); };
  if (kind === 'maintenance') return { reports: mine(noviRows_('Reports', 4), 1).slice(-15).map(function (r) { return { date: zelCell_(r[0]), issue: String(r[2]).slice(0, 300) }; }) };
  if (kind === 'closeout') return { closeouts: mine(noviRows_('Close outs', 5), 1).slice(-10).map(function (r) { return { date: zelCell_(r[0]), items_checked: r[2], notes: String(r[3]).slice(0, 300) }; }) };
  const by = {}; ordRows_().filter(function (r) { return String(r[1]) === String(loc); }).slice(-60).forEach(function (r) { const k = zelCell_(r[0]); (by[k] = by[k] || []).push('#' + r[2] + ' ' + r[3] + ' x' + r[4]); });
  return { orders: Object.keys(by).slice(-8).map(function (k) { return { sent: k, items: by[k] }; }) };
}
function zelSheet_(tab, mgr) {
  const want = String(tab || '').trim().toUpperCase(), ok = ZEL_TABS_OPEN.concat(mgr ? ZEL_TABS_MGR : []).filter(function (t) { return t.toUpperCase() === want; })[0];
  if (!ok) return { error: 'That tab is not available to you.' };
  const keys = ZEL_KEYS_OPEN.concat(ZEL_KEYS_MGR), sh = tabFind_(SpreadsheetApp.getActiveSpreadsheet(), keys[ZEL_TABS_OPEN.concat(ZEL_TABS_MGR).indexOf(ok)]); if (!sh) return { error: 'Tab not found.' };
  const n = Math.min(sh.getLastRow(), 80), w = Math.min(sh.getLastColumn(), 20);
  return { tab: ok, rows: n && w ? sh.getRange(1, 1, n, w).getValues().map(function (r) { return r.map(zelCell_); }) : [] };
}
function zelFolders_() {
  const ids = zelProp_('ZELYNA_DOC_FOLDERS');
  if (ids) return ids.split(',').map(function (x) { return x.trim(); }).filter(String);
  const out = []; try { const it = DriveApp.getFoldersByName(ZEL_SOP_FOLDER); while (it.hasNext()) out.push(it.next().getId()); } catch (e) {}
  return out;
}
// The authorized folders plus every folder inside them (up to 4 levels, 60 folders). Remembered for 10 minutes.
function zelAllFolders_() {
  const roots = zelFolders_(), ck = 'zelf:' + roots.join(',');
  try { const c = CacheService.getScriptCache().get(ck); if (c) return JSON.parse(c); } catch (e) {}
  const out = [], seen = {};
  let cur = roots.slice();
  for (let depth = 0; depth < 5 && cur.length && out.length < 60; depth++) {
    const next = [];
    cur.forEach(function (id) {
      if (seen[id] || out.length >= 60) return; seen[id] = 1; out.push(id);
      if (depth < 4) { try { const it = DriveApp.getFolderById(id).getFolders(); while (it.hasNext()) next.push(it.next().getId()); } catch (e) {} }
    });
    cur = next;
  }
  try { CacheService.getScriptCache().put(ck, JSON.stringify(out), 600); } catch (e) {}
  return out;
}
function zelInFolder_(file, ids) {
  let cur = [file];
  for (let depth = 0; depth < 6 && cur.length; depth++) {
    const next = [];
    for (let i = 0; i < cur.length; i++) { const p = cur[i].getParents(); while (p.hasNext()) { const f = p.next(); if (ids.indexOf(f.getId()) >= 0) return true; next.push(f); } }
    cur = next;
  }
  return false;
}
const ZEL_MIME_FOLDER = 'application/vnd.google-apps.folder', ZEL_MIME_GDOC = 'application/vnd.google-apps.document';
// Files Zelyna can read: Google Docs, PDFs, Word files, plain text.
function zelDocKind_(mime) {
  mime = String(mime || '');
  if (mime === ZEL_MIME_GDOC) return 'Google Doc';
  if (mime === 'application/pdf') return 'PDF';
  if (/wordprocessingml|msword|opendocument\.text|rtf/.test(mime)) return 'Word';
  if (/^text\//.test(mime)) return 'Text';
  return '';
}
function zelParentsQ_(ids) { return '(' + ids.map(function (id) { return "'" + id + "' in parents"; }).join(' or ') + ')'; }
function zelDocSearch_(q) {
  const folders = zelAllFolders_(), t = String(q || '').replace(/['\\]/g, ' ').replace(/\s+/g, ' ').trim(), seen = {}, out = [];
  if (!folders.length) return { documents: [] };
  const add = function (f) { const id = f.getId(), k = zelDocKind_(f.getMimeType()); if (!seen[id] && k && out.length < 8) { seen[id] = 1; out.push({ id: id, name: f.getName(), type: k }); } };
  const words = t.split(' ').filter(function (w) { return w.length > 3; });
  const tries = t ? [t].concat(words.length > 1 ? words : []) : [];
  for (let c = 0; c < folders.length; c += 20) {
    const pq = zelParentsQ_(folders.slice(c, c + 20));
    tries.forEach(function (term) {
      ['title contains', 'fullText contains'].forEach(function (op) {
        if (out.length >= 8) return;
        try { const it = DriveApp.searchFiles(pq + " and trashed = false and mimeType != '" + ZEL_MIME_FOLDER + "' and " + op + " '" + term + "'"); while (it.hasNext() && out.length < 8) add(it.next()); } catch (e) {}
      });
    });
  }
  if (out.length) return { documents: out };
  return { documents: [], available: zelDocList_(25), note: 'No document matched. These are the documents you may read; pick one by its id.' };
}
// Every readable document in the authorized folders (names only).
function zelDocList_(max) {
  const out = [];
  zelAllFolders_().forEach(function (fid) {
    if (out.length >= max) return;
    try { const it = DriveApp.getFolderById(fid).getFiles(); while (it.hasNext() && out.length < max) { const f = it.next(), k = zelDocKind_(f.getMimeType()); if (k) out.push({ id: f.getId(), name: f.getName(), type: k }); } } catch (e) {}
  });
  return out;
}
function zelGdocText_(id) {
  const doc = DocumentApp.openById(id), parts = [];
  const walk = function (tabs) { tabs.forEach(function (tb) { try { parts.push('[' + tb.getTitle() + ']\n' + tb.asDocumentTab().getBody().getText()); } catch (e) {} walk(tb.getChildTabs ? tb.getChildTabs() : []); }); };
  try { walk(doc.getTabs()); } catch (e) {}
  return parts.length ? parts.join('\n\n') : doc.getBody().getText();
}
// PDFs and Word files: Google Drive turns a temporary copy into a Google Doc (it reads scanned pages too), Zelyna reads it,
// and the copy is deleted right away. The text is remembered until the file changes, so each file is converted only once.
function zelConvText_(f) {
  const ver = String(f.getLastUpdated().getTime()), ck = 'zelpdf:' + f.getId() + ':' + ver;
  try { const c = CacheService.getScriptCache().get(ck); if (c !== null) return c; } catch (e) {}
  const store = PropertiesService.getScriptProperties(), pk = 'ZPDF_' + f.getId();
  try { const m = JSON.parse(store.getProperty(pk) || 'null'); if (m && m.v === ver && m.file) { const tx = DriveApp.getFileById(m.file).getBlob().getDataAsString('UTF-8'); try { CacheService.getScriptCache().put(ck, tx.slice(0, 90000), 21600); } catch (e) {} return tx; } } catch (e) {}
  const tok = ScriptApp.getOAuthToken(), api = 'https://www.googleapis.com/drive/v3/files/';
  const r = UrlFetchApp.fetch(api + encodeURIComponent(f.getId()) + '/copy?supportsAllDrives=true&fields=id', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + tok },
    payload: JSON.stringify({ name: 'zelyna-reading-' + f.getName(), mimeType: ZEL_MIME_GDOC })
  });
  if (r.getResponseCode() >= 300) throw new Error('convert ' + r.getResponseCode() + ' ' + String(r.getContentText()).slice(0, 160));
  const tmp = JSON.parse(r.getContentText()).id;
  let text = '';
  try { text = zelGdocText_(tmp); }
  finally {
    try { UrlFetchApp.fetch(api + encodeURIComponent(tmp) + '?supportsAllDrives=true', { method: 'delete', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + tok } }); } catch (e) {}
    try { DriveApp.getFileById(tmp).setTrashed(true); } catch (e) {}
  }
  text = String(text || '').replace(/\n{3,}/g, '\n\n').trim();
  try { CacheService.getScriptCache().put(ck, text.slice(0, 90000), 21600); } catch (e) {}
  // keep a small text copy in the data folder so the next reading is instant
  try {
    const old = JSON.parse(store.getProperty(pk) || 'null'); if (old && old.file) { try { DriveApp.getFileById(old.file).setTrashed(true); } catch (e) {} }
    const tf = logFolder_().createFile('zelyna-text-' + f.getId() + '.txt', text, 'text/plain');
    store.setProperty(pk, JSON.stringify({ v: ver, file: tf.getId() }));
  } catch (e) {}
  return text;
}
function zelDocRead_(id) {
  const ids = zelFolders_(); let f;
  try { f = DriveApp.getFileById(String(id || '')); } catch (e) { return { error: 'Document not found.' }; }
  if (!zelInFolder_(f, ids)) return { error: 'That document is not in the authorized folder.' };
  const kind = zelDocKind_(f.getMimeType());
  let text = '';
  try {
    if (kind === 'Google Doc') text = zelGdocText_(f.getId());
    else if (kind === 'PDF' || kind === 'Word') text = zelConvText_(f);
    else if (kind === 'Text') text = f.getBlob().getDataAsString('UTF-8');
    else return { error: 'Zelyna can read Google Docs, PDFs and Word files only. "' + f.getName() + '" is another kind of file.' };
  } catch (e) {
    console.log('zelDocRead_ ' + f.getName() + ': ' + e);
    return { error: 'Could not read "' + f.getName() + '". If it is a very large PDF, open it in Google Drive and choose Open with > Google Docs, then keep that Google Doc in the folder.' };
  }
  if (!String(text).trim()) return { name: f.getName(), type: kind, text: '', error: 'This file has no readable text (it may be only pictures).' };
  return { name: f.getName(), type: kind, text: text.slice(0, 12000), truncated: text.length > 12000 };
}
function zelRun_(name, input, ctx) {
  input = input || {};
  try {
    if (name === 'inventory_lookup') return zelInventory_(ctx.loc, input.query);
    if (name === 'temperature_status') return zelTemps_();
    if (name === 'task_status') return ctx.mgr ? zelTasks_(input.employee) : { error: 'Manager mode is required for employee and task information. Ask the person to turn on Manager mode in the Collaborator task list.' };
    if (name === 'recent_records') return zelRecords_(ctx.loc, input.kind);
    if (name === 'read_sheet') return zelSheet_(input.tab, ctx.mgr);
    if (name === 'get_recipe') return rcTool_(input && input.name);
    if (name === 'search_knowledge') return zkbSearch_(input.query, input.category);
    if (name === 'search_documents') return zelDocSearch_(input.query);
    if (name === 'read_document') return zelDocRead_(input.id);
    if (name === 'log_unanswered_question') return zelFaqLog_(input.question, input.category);
  } catch (e) { return { error: 'Lookup failed: ' + String(e).slice(0, 120) }; }
  return { error: 'Unknown tool.' };
}
const ZEL_FAQ_TAB = 'ZELYNA Q&A';
// Common questions and the answers Zelyna must use first. Edit the tab in the Sheet any time; no code change needed.
const ZEL_FAQ_HEAD = ['Question or comment', 'Answer (leave empty for a comment or rule)', 'Active', 'Category (filled in automatically)', 'Status (filled in automatically)'];
function zelFaqSetup_(ss, out) {
  let sh = tabFind_(ss, ZEL_FAQ_TAB);
  if (sh && sh.getLastRow() >= 2) {
    if (String(sh.getRange(1, 4).getValue() || '') !== ZEL_FAQ_HEAD[3]) { sh.getRange(1, 4, 1, 2).setValues([ZEL_FAQ_HEAD.slice(3)]).setFontWeight('bold'); try { sh.setColumnWidth(4, 170); sh.setColumnWidth(5, 320); } catch (e) {} }
    try { zelFaqTidy_(); } catch (e) {}
    out.push('ZELYNA Q&A tab already has rows, left as is (Category and Status columns ready).'); return;
  }
  if (!sh) sh = ss.insertSheet(tabName_(ZEL_FAQ_TAB));
  const rows = [
    ['What is the walk-in cooler temperature limit?', 'The walk-in must stay at 41 F or colder. If it is above 41 F, tell the manager right away.', 'Yes', 'Food safety', ''],
    ['How long can cooked food stay in the danger zone?', 'No more than 2 hours between 41 F and 135 F. Cool it fast or throw it out.', 'Yes', 'Food safety', ''],
    ['Where do I find the cleaning checklist?', 'Open the Today task list on the home screen. Each cleaning item is listed there.', 'Yes', 'Cleaning', ''],
    ['What do I do when an item is out of stock?', 'Mark it OUT in Inventory so it shows on the next order, and tell the manager.', 'Yes', 'Inventory and orders', ''],
    ['(Type each common question in column A and its answer in column B. For a general comment or rule, type it in column A and leave B empty. Questions Zelyna could not answer are added here automatically: type the answer in column B.)', '', 'No', '', ''],
    ['Always remind the cashier to check the dining area at the times on the Alarms tab.', '', 'No', 'Customer service', '']
  ];
  sh.getRange(1, 1, 1, 5).setValues([ZEL_FAQ_HEAD]).setFontWeight('bold').setBackground('#F4B400');
  sh.getRange(2, 1, rows.length, 5).setValues(rows);
  try { sh.setFrozenRows(1); sh.setColumnWidth(1, 380); sh.setColumnWidth(2, 560); sh.setColumnWidth(3, 70); sh.setColumnWidth(4, 170); sh.setColumnWidth(5, 320); sh.getRange(2, 1, 400, 2).setWrap(true); sh.getRange(2, 3, 400, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build()); } catch (e) {}
  try { tabLook_(sh); } catch (e) {}
  try { zelFaqTidy_(); } catch (e) {}
  out.push('ZELYNA Q&A tab created (' + rows.length + ' sample rows).');
}
function zelFaqSheet_() { return tabFind_(SpreadsheetApp.getActiveSpreadsheet(), ZEL_FAQ_TAB); }
function zelFaqRows_() { const sh = zelFaqSheet_(); if (!sh || sh.getLastRow() < 2) return []; return sh.getRange(2, 1, Math.min(sh.getLastRow() - 1, 400), 5).getValues(); }
function zelOn_(v) { return String(v === null || v === undefined ? '' : v).trim().toLowerCase() !== 'no'; }
function zelFaqNorm_(q) { return String(q || '').toLowerCase().replace(/[¿?¡!.,;:"']/g, ' ').replace(/\s+/g, ' ').trim(); }
// Approved answers + owner notes for the system prompt, grouped by category. Questions waiting for an answer are left out.
function zelFaq_() {
  try {
    const by = {};
    zelFaqRows_().forEach(function (r) {
      const q = String(r[0] || '').trim(), a = String(r[1] || '').trim(); if (!q || !zelOn_(r[2])) return;
      const c = String(r[3] || '').trim() || 'Other';
      (by[c] = by[c] || []).push(a ? 'Q: ' + q + '\nA: ' + a : 'NOTE FROM THE OWNER: ' + q);
    });
    return Object.keys(by).map(function (c) { return '[' + c + ']\n' + by[c].join('\n'); }).join('\n').slice(0, 9000);
  } catch (e) { return ''; }
}
// Zelyna could not answer: add the question (or count it again) so the owner can type the answer.
function zelFaqLog_(question, cat) {
  const q = String(question || '').trim().slice(0, 300); if (q.length < 4) return { ok: false };
  const sh = zelFaqSheet_(); if (!sh) return { ok: false, error: 'No ZELYNA Q&A tab yet.' };
  const rows = zelFaqRows_(), k = zelFaqNorm_(q), day = logFmt_(logNow_().date);
  for (let i = 0; i < rows.length; i++) {
    if (zelFaqNorm_(rows[i][0]) === k) {
      const m = /asked (\d+)/.exec(String(rows[i][4] || '')), n = (m ? +m[1] : 1) + 1;
      if (!String(rows[i][1] || '').trim()) sh.getRange(i + 2, 5).setValue('❓ Needs an answer · asked ' + n + ' times · last ' + day);
      return { ok: true, saved: 'counted again' };
    }
  }
  const c = ZEL_CATS.indexOf(String(cat || '')) >= 0 ? cat : zelCatGuess_(q);
  sh.getRange(Math.max(sh.getLastRow(), 1) + 1, 1, 1, 5).setValues([[q, '', 'No', c, '❓ Needs an answer · asked 1 time · last ' + day]]);
  return { ok: true, saved: 'added for the owner' };
}
function zelCatGuess_(q) {
  const t = String(q || '').toLowerCase(), R = [
    ['Food safety', /temp|°|degree|grado|cool|cold|frío|frio|hot|caliente|thaw|descongel|danger|peligro|allerg|alerg|glove|guante|hand ?wash|lavar las manos|label|etiqueta|fifo|expire|caduc|use by|shelf/],
    ['Cleaning', /clean|limpi|sanit|wash|lavar|mop|trapear|sweep|barrer|trash|basura|bleach|cloro/],
    ['Equipment', /fryer|freidora|grill|plancha|oven|horno|fridge|refri|cooler|walk-in|machine|máquina|maquina|printer|impresora|tablet|broken|roto|repair|repar|leak|fuga/],
    ['Inventory and orders', /inventor|order|orden|pedido|stock|out of|falta|supplier|proveedor|delivery|entrega/],
    ['Prep and recipes', /prep|recipe|receta|marina|cook|cocin|salsa|rice|arroz|bean|frijol|meat|carne|asada|pollo|chicken|pastor|birria|cut|picar|portion|porci/],
    ['Cash and register', /cash|caja|register|pos|card|tarjeta|tip|propina|refund|reembolso|change|cambio|drawer/],
    ['Customer service', /customer|cliente|complain|queja|order wrong|dining|comedor|table|mesa|phone|teléfono|telefono/],
    ['Schedule and team', /schedule|horario|shift|turno|break|descanso|late|tarde|sick|enferm|pay|pago|uniform|uniforme|day off/],
    ['Opening and closing', /open|abrir|close|cerrar|closing|cierre|lock|alarm|alarma/]];
  for (let i = 0; i < R.length; i++) if (R[i][1].test(t)) return R[i][0];
  return 'Other';
}
// Fills Category (AI when the key is set, otherwise keywords) and Status for every row. Runs on every edit of the tab.
function zelFaqTidy_() {
  const sh = zelFaqSheet_(); if (!sh || sh.getLastRow() < 2) return { ok: true, rows: 0 };
  const rows = zelFaqRows_(), need = [];
  rows.forEach(function (r, i) { if (String(r[0] || '').trim() && !String(r[3] || '').trim()) need.push(i); });
  let cats = {};
  if (need.length && zelKey_()) { try { cats = zelCatAi_(need.map(function (i) { return String(rows[i][0]) + (rows[i][1] ? ' — ' + String(rows[i][1]).slice(0, 160) : ''); })); } catch (e) {} }
  const C = [], St = [], act = [];
  rows.forEach(function (r, i) {
    const q = String(r[0] || '').trim(), a = String(r[1] || '').trim(), st0 = String(r[4] || '');
    let on = zelOn_(r[2]), c = String(r[3] || '').trim();
    if (q && !c) { const k = need.indexOf(i); c = (cats[k] && ZEL_CATS.indexOf(cats[k]) >= 0) ? cats[k] : zelCatGuess_(q); }
    if (q && a && /Needs an answer/.test(st0) && !on) on = true;   // the owner typed the answer to a logged question: use it
    act.push([q ? (on ? 'Yes' : 'No') : r[2]]);
    C.push([q ? c : r[3]]);
    St.push([!q ? '' : /^\(/.test(q) ? '' : !on ? (/Needs an answer/.test(st0) && !a ? st0 : 'Hidden (Active = No)') : a ? '✓ Zelyna uses this answer' : '✓ Zelyna follows this note as a rule']);
  });
  sh.getRange(2, 3, rows.length, 1).setValues(act);
  sh.getRange(2, 4, rows.length, 1).setValues(C);
  sh.getRange(2, 5, rows.length, 1).setValues(St);
  return { ok: true, rows: rows.length };
}
function zelCatAi_(list) {
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': zelKey_(), 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: zelModel_(), max_tokens: 800, system: 'You sort restaurant staff questions into categories. Reply with ONLY a JSON array of category names, one per question, in the same order. Allowed categories: ' + ZEL_CATS.join(', ') + '.',
      messages: [{ role: 'user', content: list.map(function (q, i) { return (i + 1) + '. ' + q; }).join('\n') }] }) });
  const j = JSON.parse(res.getContentText()), txt = (j.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
  const arr = JSON.parse(txt.slice(txt.indexOf('['), txt.lastIndexOf(']') + 1)), out = {};
  arr.forEach(function (c, i) { out[i] = String(c); });
  return out;
}
// Returns the system prompt as two blocks: the stable part (cached by the AI service, so later questions start faster) and the live part.
function zelSystem_(loc, lang, mgr) {
  const now = logNow_(), faq = zelFaq_();
  const stable = 'You are Zelyna, the assistant inside Digital Pilot OPS for the restaurant location "' + loc + '". Kitchen staff and managers talk to you by voice or text on a tablet. ' +
    'HOW TO ANSWER: Be short and direct. Prefer "Yes", "No", "Done" or one short sentence. Give details, lists or steps only when the person asks for them; when steps are needed, give only the needed steps. ' +
    'Never greet, never introduce yourself, never repeat the question, never add a closing line. Sound natural, like a coworker; plain sentences that read well aloud; no markdown, no tables, no emojis. ' +
    'LANGUAGE: Answer in the language of the person\'s latest message. Spanish in, Spanish out; English in, English out. If the message mixes Spanish and English, answer in the language most of it uses. Keep names, Mexican Spanish words and kitchen terms exactly as said (asada, pastor, birria, picar, carnitas...). Translate only when asked. ' +
    'CONVERSATION: Use the earlier messages to understand short follow-ups ("and the chicken?", "¿y el otro?") without asking the person to repeat. ' +
    'DATA: Use the tools to look things up and answer ONLY from what they return or from the approved Q&A. Never invent numbers, names, temperatures, inventory or document content. If the information is not available, say so in one sentence. ' +
    'Give ONE answer. When several Q&A rows, documents or records match, use the single best one (the most specific; when two Q&A rows disagree, the lower row is the newer one). Never list several possible answers. Ask one short question back only when the request truly cannot be answered without it. ' +
    'LIMITS: You are an assistant, not a manager. You only answer questions and read authorized information. Do not recommend actions or solutions, do not assign or suggest tasks, do not draft inventory orders, do not approve or reject anything, and you cannot change anything in the app. If asked to do one of these, say in one sentence that a manager does it in the app. When the data shows a problem, state the fact plainly without advice. ' +
    'PHOTOS: a search result or recipe can have photo_ids. When the person asks how something looks, how to plate, build or pack it, or a photo clearly helps, end your reply with [photo:ID] (the id exactly as given, at most 2). Never write an id that was not given to you. ' +
    'WHERE TO LOOK (in this order, stop as soon as you have the answer): 1) the approved Q&A below; 2) search_knowledge, with the category of the question (food safety, cleaning, SOPs, daily operations, training, menu for how a menu item is made, built, plated or packed); for the ingredients and amounts of a batch recipe use get_recipe (the RECIPES tab); 3) read_document only when the section found is not enough; 4) search_documents only when the first steps found nothing (for example a PDF). Do not search every document for every question. Keep answers short, but always keep the safety facts of the answer: temperatures, times, and what must never be done. ' +
    'If you cannot find the answer in the approved Q&A, the tools or the SOP documents, call log_unanswered_question once, then say you do not know yet and that the question was saved for the manager.' +
    (faq ? '\n\nAPPROVED Q&A and NOTES from the owner (use these answers first; follow every NOTE as a rule):\n' + faq : '');
  const live = 'Today is ' + now.date + ' and the time is ' + noviTime_() + '. The app is set to ' + (lang === 'es' ? 'Spanish' : 'English') + ' (only a hint; answer in the language of the latest message). Manager mode is ' + (mgr ? 'ON (employee and task information is allowed)' : 'OFF (do not discuss individual employees; say Manager mode is needed in the Collaborator task list)') + '.';
  return [{ type: 'text', text: stable, cache_control: { type: 'ephemeral' } }, { type: 'text', text: live }];
}
// Voice: the tablet sends the recording; it is turned into text here (OpenAI speech-to-text, English/Spanish detected
// automatically, nothing translated) and answered in the same request, so there is only one wait.
const ZEL_STT_HINT = 'Restaurant kitchen at Birriaz, a Mexican restaurant. People speak English, Spanish or both mixed. Write exactly what is said in the language it is said; do not translate. Words: Zelyna, asada, pollo, al pastor, birria, carnitas, quesabirria, cebolla, cilantro, repollo, lechuga, pico de gallo, guacamole, salsa roja, salsa verde, frijoles, arroz, marinar, picar, walk-in, FIFO, inventario, tareas.';
function zelStt_(b64, mime, lang) {
  const say = function (en, es) { return lang === 'es' ? es : en; };
  const key = zelProp_('OPENAI_API_KEY');
  if (!key) return { ok: false, code: 'no_stt', error: say('Voice is not set up yet. Ask the owner to add the OPENAI_API_KEY in Script Properties.', 'La voz no está configurada. Pida al dueño que agregue OPENAI_API_KEY en Script Properties.') };
  const m = /^audio\/[a-z0-9.+-]+/i.test(String(mime || '')) ? String(mime).split(';')[0] : 'audio/webm', ext = { 'audio/webm': 'webm', 'audio/mp4': 'm4a', 'audio/ogg': 'ogg', 'audio/mpeg': 'mp3', 'audio/wav': 'wav' }[m] || 'webm';
  let res, j;
  try {
    res = UrlFetchApp.fetch('https://api.openai.com/v1/audio/transcriptions', { method: 'post', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + key },
      payload: { file: Utilities.newBlob(Utilities.base64Decode(String(b64)), m, 'speech.' + ext), model: zelProp_('OPENAI_STT_MODEL') || 'whisper-1', prompt: ZEL_STT_HINT, response_format: 'json', temperature: '0' } });
    j = JSON.parse(res.getContentText());
  } catch (e) { Logger.log('STT failed: ' + e); return { ok: false, code: 'stt', error: say('I could not hear the recording right now. Please try again.', 'No pude escuchar la grabación ahora. Intente de nuevo.') }; }
  if (res.getResponseCode() !== 200 || !j || typeof j.text !== 'string') { Logger.log('STT error ' + res.getResponseCode() + ' ' + String(res.getContentText()).slice(0, 200)); return { ok: false, code: 'stt', error: say('I could not hear the recording right now. Please try again.', 'No pude escuchar la grabación ahora. Intente de nuevo.') }; }
  let text = j.text.replace(/\s+/g, ' ').trim();
  if (/^(thank you( for watching)?|thanks for watching|gracias( por ver)?|subt[ií]tulos.*|you|\.+)[.!]?$/i.test(text)) text = '';   // what the service writes for silence
  return { ok: true, text: text.slice(0, 600) };
}
function zelynaAsk_(b) {
  b = b || {};
  if (!b.audio) return zelynaAsk0_(b);
  const lang = b.lang === 'es' ? 'es' : 'en';
  if (!zelKey_() || !zelLocOk_(String(b.location || '').trim().slice(0, 50))) return zelynaAsk0_(Object.assign({}, b, { audio: null }));
  if (String(b.audio).length > 8000000) return { ok: false, error: lang === 'es' ? 'La grabación es muy larga.' : 'The recording is too long.' };
  const tr = zelStt_(b.audio, b.mime, lang); if (!tr.ok) return tr;
  if (!tr.text) return { ok: false, code: 'empty', error: lang === 'es' ? 'No escuché nada, así que no se envió nada. Toque el micrófono e intente de nuevo.' : "I didn't hear anything, so nothing was sent. Tap the microphone and try again." };
  const r = zelynaAsk0_(Object.assign({}, b, { audio: null, messages: (Array.isArray(b.messages) ? b.messages : []).concat([{ role: 'user', content: tr.text }]) }));
  r.heard = tr.text; return r;
}
function zelynaAsk0_(b) {
  b = b || {};
  const loc = String(b.location || '').trim().slice(0, 50), lang = b.lang === 'es' ? 'es' : 'en';
  const say = function (en, es) { return lang === 'es' ? es : en; };
  const key = zelKey_();
  if (!key) return { ok: false, code: 'no_key', error: say('Zelyna is not connected to the AI yet. Ask the owner to add the Zelyna AI key.', 'Zelyna todavía no está conectada a la IA. Pida al dueño que agregue la clave de Zelyna.') };
  if (!zelLocOk_(loc)) return { ok: false, code: 'bad_location', error: say('I do not know this location.', 'No conozco esta ubicación.') };
  let msgs = (Array.isArray(b.messages) ? b.messages : []).slice(-12).map(function (m) { return { role: m && m.role === 'assistant' ? 'assistant' : 'user', content: String((m && m.content) || '').slice(0, 1500) }; }).filter(function (m) { return m.content.trim(); });
  while (msgs.length && msgs[0].role !== 'user') msgs.shift();
  if (!msgs.length || msgs[msgs.length - 1].role !== 'user') return { ok: false, error: say('Ask me a question.', 'Hágame una pregunta.') };
  if (!zelCap_(loc)) return { ok: false, code: 'limit', error: say('I reached today\'s question limit for this location. Please try again tomorrow.', 'Llegué al límite de preguntas de hoy para esta ubicación. Intente mañana.') };
  const ctx = { loc: loc, mgr: String(b.pin || '') === TW_PIN }, sys = zelSystem_(loc, lang, ctx.mgr);
  const q = String(msgs[msgs.length - 1].content || ''), fast = zelFastModel_(q, msgs, ctx), strong = zelModel_();
  let r = null, model = strong;
  if (fast && fast !== strong) { r = zelLoop_(key, fast, sys, msgs.slice(), ctx, say, true); model = fast; if (r.escalate) { r = null; model = strong; } }   // simple question: the fast model first
  if (!r) r = zelLoop_(key, strong, sys, msgs.slice(), ctx, say, false);
  if (!r.ok) return { ok: false, error: r.error };
  try { zqLog_(q, r.used.indexOf('log_unanswered_question') < 0, lang, model); } catch (x) {}
  // [photo:ID] at the end of the reply = show that photo (only photos Zelyna was given; the marks are taken out of the text)
  const photos = [];
  let reply = String(r.reply || '').replace(/\[photo:\s*([-\w]{20,80})\s*\]/gi, function (m, id) { if (photos.length < 2 && photos.indexOf(id) < 0 && zphOk_(id)) photos.push(id); return ''; }).replace(/[ \t]+\n/g, '\n').trim();
  const out = { ok: true, reply: reply || say('I do not have an answer for that.', 'No tengo una respuesta para eso.'), tools: r.used, model: model };
  if (photos.length) out.photos = photos;
  return out;
}
// Faster answers for simple questions: short lookups go to the fast model (Script Property ZELYNA_FAST_MODEL, default claude-haiku-5-5).
// Longer questions, or ones that ask to explain / compare / summarize / look at reports or weeks, stay on the main model.
// ZELYNA_FAST = off turns it off. If the fast model cannot answer, the same question goes to the main model by itself.
const ZEL_FAST_DEFAULT = 'claude-haiku-5-5';
const ZEL_HARD = /\b(why|explain|compare|comparison|summar\w*|report|reports|trend|analy\w*|weekly|this week|last week|month|plan|schedule|difference|porque|por qu[eé]|expl[ií]ca\w*|compar\w*|resum\w*|reporte\w*|tendencia|anali\w*|semana\w*|mes|diferencia)(?=$|[^\wáéíóúñ])/i;
function zelFastModel_(q, msgs, ctx) {
  if (String(zelProp_('ZELYNA_FAST') || '').toLowerCase() === 'off') return '';
  q = String(q || '').trim();
  const words = q.split(/\s+/).filter(String).length, qs = (q.match(/\?/g) || []).length;
  if (!q || q.length > 140 || words > 20 || qs > 1 || ZEL_HARD.test(q)) return '';
  return zelProp_('ZELYNA_FAST_MODEL') || ZEL_FAST_DEFAULT;
}
// One answer with one model (the tools / knowledge search are the same for both). trial = the fast attempt: it never saves
// a question as unanswered; it gives the question to the main model instead (also when it fails or has no answer).
const ZEL_NOANS = /\b(i (do not|don't) (know|have)|i could( not|n't) find|not sure|no (lo )?s[eé]|no tengo|no encontr[eé]|no estoy segur)/i;
function zelLoop_(key, model, sys, msgs, ctx, say, trial) {
  const used = [];
  for (let step = 0; step < 6; step++) {
    let res, j;
    try {
      res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: model, max_tokens: 600, system: sys, tools: ZEL_TOOLS, messages: msgs }) });
      j = JSON.parse(res.getContentText());
    } catch (e) { Logger.log('Zelyna fetch failed (' + model + '): ' + e); return trial ? { escalate: true } : { ok: false, error: say('I could not reach my AI right now. Please try again.', 'No pude conectarme a mi IA ahora. Intente de nuevo.') }; }
    if (res.getResponseCode() !== 200 || !j || !Array.isArray(j.content)) { Logger.log('Zelyna API error (' + model + ') ' + res.getResponseCode() + ' ' + String(res.getContentText()).slice(0, 200)); return trial ? { escalate: true } : { ok: false, error: say('I could not answer right now. Please try again.', 'No pude responder ahora. Intente de nuevo.') }; }
    const calls = j.content.filter(function (c) { return c.type === 'tool_use'; });
    if (j.stop_reason === 'tool_use' && calls.length) {
      if (trial && calls.some(function (c) { return c.name === 'log_unanswered_question'; })) return { escalate: true };
      msgs.push({ role: 'assistant', content: j.content });
      msgs.push({ role: 'user', content: calls.map(function (c) { used.push(c.name); return { type: 'tool_result', tool_use_id: c.id, content: JSON.stringify(zelRun_(c.name, c.input, ctx)).slice(0, 14000) }; }) });
      continue;
    }
    const reply = j.content.filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('\n').trim();
    if (trial && (!reply || ZEL_NOANS.test(reply))) return { escalate: true };
    return { ok: true, reply: reply, used: used };
  }
  return trial ? { escalate: true } : { ok: false, error: say('That took too long. Try a simpler question.', 'Tardó demasiado. Intente una pregunta más simple.') };
}
// Run once from the editor to check the connection (writes the answer to the log).
function zelynaTest() { const r = zelynaAsk_({ location: zelLocs_()[0] || 'San Jacinto', lang: 'en', messages: [{ role: 'user', content: 'Hi Zelyna, what is marked out in inventory right now?' }] }); Logger.log(JSON.stringify(r)); return r; }


/* ===================== MAINTENANCE ISSUES (Collaborator task list -> Today's task list) =====================
   One JSON file (maint-issues.json) in the same data folder holds every issue. Photos are saved as small JPG files in the same folder.
   An issue stays on Today's task list every day until someone marks it completed (that is the "roll over"). */
const MR_FILE = 'maint-issues.json';
function mrRead_() { const d = logReadJson_(MR_FILE); return (d && Array.isArray(d.items)) ? d : { items: [] }; }
function mrPublic_(it) { return { id: it.id, text: it.text, emp: it.emp || '', em: !!it.em, photo: !!it.photo, by: it.by || '', cts: it.cts || '', done: !!it.done, doneAt: it.doneAt || '', doneBy: it.doneBy || '' }; }
function mrList_() {
  const d = mrRead_(), today = logNow_().date, keep = logAddDays_(today, -2);
  const items = d.items.filter(function (it) { return !it.removed && (!it.done || String(it.doneAt || '').slice(0, 10) >= keep); });
  return { ok: true, today: today, items: items.map(mrPublic_) };
}
// COLLAB REPORTS (filled in by the system): every check a collaborator makes in their own profile on a tablet, with the time.
// Result = Completed / Undone / Blocked: too soon (the profile allows one task per minute). Seconds = time since their last completed task.
const CR_HEAD = ['Date & time', 'Collaborator', 'Task', 'Kind', 'Result', 'Seconds since their last task', 'Location', 'Entry ID (do not change)'];
function collabRep_(raw) {
  const list = (raw && Array.isArray(raw.items) ? raw.items : []).slice(0, 50);
  if (!list.length) return { ok: true, n: 0 };
  const sh = getTab_('COLLAB REPORTS', CR_HEAD), last = sh.getLastRow(), seen = {};
  let c = null; try { c = CacheService.getScriptCache(); } catch (e) {}
  if (last > 1) { const from = Math.max(2, last - 499); sh.getRange(from, 8, last - from + 1, 1).getValues().forEach(function (r) { seen[String(r[0])] = 1; }); }
  const rows = [];
  list.forEach(function (x) {
    const id = String((x && x.id) || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
    if (!id || seen[id] || (c && c.get('cr_' + id))) return;
    seen[id] = 1; try { if (c) c.put('cr_' + id, '1', 21600); } catch (e) {}
    const at = Number(x.at), d = isFinite(at) && at > 0 ? new Date(at) : new Date();
    const gap = x.gap === '' || x.gap === null || x.gap === undefined || !isFinite(Number(x.gap)) ? '' : Math.max(0, Math.round(Number(x.gap)));
    rows.push([d, String(x.emp || '').slice(0, 40), String(x.task || '').slice(0, 200), String(x.kind || '').slice(0, 30), String(x.result || '').slice(0, 60), gap, String(x.location || '').slice(0, 40), id]);
  });
  if (rows.length) {
    const r = sh.getRange(sh.getLastRow() + 1, 1, rows.length, CR_HEAD.length);
    r.setValues(rows);
    try { sh.getRange(r.getRow(), 1, rows.length, 1).setNumberFormat('ddd mmm d, yyyy h:mm:ss am/pm'); } catch (e) {}
  }
  return { ok: true, n: rows.length };
}
/* ===================== RECIPES tab: the Recipes screen on the tablets =====================
   One row per line. Recipe = which recipe; rows of one recipe can be anywhere, the app shows the recipes and their lines in the order of the tab.
   A row with only Notes = a written line (a step); when it is the first row of a recipe it is shown on top (container size).
   Active = No hides the row; a recipe with no active rows is hidden. The tablets keep the last list they got. */
const RC_TAB = 'RECIPES', RC_HEAD = ['Recipe', 'Ingredient', 'Measurement', 'Amount', 'Notes', 'Active', 'Status (filled in automatically)', 'Photo (paste a Google Drive link, on any row of the recipe)'];
const RC_SEED = [["CEBOLLA MORADA", "CEBOLLA", "18 QT", "SLICED", ""], ["CEBOLLA MORADA", "LIMON JUICE", "OZ", "6", ""], ["CEBOLLA MORADA", "HABANERO", "PCS", "25", ""], ["CEBOLLA MORADA", "SAL", "TBSP", "3", ""], ["CEBOLLA MORADA", "BLK PEPPER", "TBSP", "1", ""], ["CEBOLLA MORADA", "VINAGRE", "ML", "50", ""], ["PICO", "TOMATE", "LBS", "4", ""], ["PICO", "LIMON JICE", "OZ", "4", ""], ["PICO", "CEBOLLA", "OZ", "24", ""], ["PICO", "CILANTRO", "OZ", "24", ""], ["PICO", "SAL", "TSP", "1", ""], ["PICO", "AJO", "TSP", "1/2", ""], ["GUAC SAUCE", "AGUACATE", "OZ", "21", ""], ["GUAC SAUCE", "LIME JUICE", "OZ", "6", ""], ["GUAC SAUCE", "SAL DE AJO", "1TBSP + 2 TSP", "", ""], ["GUAC SAUCE", "WATER", "OZ", "40", ""], ["GUACAMOLE", "AGUACATE", "6 LBS -2.5 OZ", "", ""], ["GUACAMOLE", "LIMON JUICE", "OZ", "8", ""], ["GUACAMOLE", "SAL DE AJO", "TBSP 2 + 1 TSP", "", ""], ["ARROZ BLANCO", "ARROZ", "QT", "4", ""], ["ARROZ BLANCO", "AGUA", "QT", "4", ""], ["ARROZ BLANCO", "LAUREL", "PZA", "2", ""], ["ARROZ BLANCO", "SAL", "TBSP", "2", ""], ["ARROZ BLANCO", "", "", "", "separe en 2 recipientes transparentes. Agregue 1 de aceite por recepiente"], ["FAJITAS", "GREEN BELL", "PCS", "6", ""], ["FAJITAS", "RED BEEL", "PCS", "1", ""], ["FAJITAS", "YELLOW BELL", "PCS", "1", ""], ["FAJITAS", "WHITE ONION", "PCS", "2", ""], ["BEANS", "BEANS", "QT", "24", ""], ["BEANS", "WATER", "FILL", "", ""], ["BEANS", "SALT", "SPOONS", "3", ""], ["PINK SAUCE", "MAYO KEN'S", "GALLON", "1", ""], ["PINK SAUCE", "POWDER", "BAG", "1", ""], ["PINK SAUCE", "WATER", "WATER", "32", ""], ["OIL", "OIL", "CONT", "3", ""], ["OIL", "PAP", "CONT", "2", ""], ["ARROZ ROJO", "TOMATE", "LBS", "7.5", ""], ["ARROZ ROJO", "CB", "PCS", "2", ""], ["ARROZ ROJO", "AJO", "PCS", "4", ""], ["ARROZ ROJO", "CM", "OZ", "12", ""], ["ARROZ ROJO", "AGUA", "LT", "3", ""], ["ARROZ ROJO", "ARROZ ROJO", "QT", "4", ""], ["HORCHATA", "", "", "", "8 QT RED CONTAINER"], ["HORCHATA", "RICE", "SPOONS", "2", ""], ["HORCHATA", "", "", "", "HEAT 1 CHOCOLATE AND 4 STICKS OF CINNAMON, 1:30 IN WATER USING A 2QT CONTAINER BLEND AND (DO NOT STRAIN)"], ["HORCHATA", "HORCHATA SYRUP", "OZ", "48", ""], ["HORCHATA", "", "", "", "WATER, FILL TO 7.5 QT"], ["MAZAPAN HORCHATA", "", "", "", "8 QT RED CONTAINER"], ["MAZAPAN HORCHATA", "RICE", "SPOONS", "3", ""], ["MAZAPAN HORCHATA", "", "", "", "STRAIN RICE"], ["MAZAPAN HORCHATA", "CINNAMON STICS", "PCS", "4", ""], ["MAZAPAN HORCHATA", "MAZAPAN POWDER", "OZ", "16", ""], ["MAZAPAN HORCHATA", "EVAPORATED MILK", "CUPS", "4", ""], ["MAZAPAN HORCHATA", "SUGAR", "CUPS", "1-1/2", ""], ["MAZAPAN HORCHATA", "CONDENSE MILK", "CANS", "2", ""], ["MAZAPAN HORCHATA", "VANILLA EXTRACT", "TBSP", "2", ""], ["Cinnamon Sugar", "sugar", "spoons", "3", ""], ["Cinnamon Sugar", "cinnamon", "Tbsp", "3", ""], ["BUBBLER", "", "", "", "8 QT RED CONTAINER"], ["BUBBLER", "SYRUP", "GAL", "1/2", ""], ["BUBBLER", "SUGAR", "CUP", "1/2", ""], ["BUBBLER", "", "", "", "FILL WATER TO 6 QT"]];
function rcSetup_(ss) {
  let sh = tabFind_(ss, RC_TAB); if (sh && sh.getLastRow() >= 2) { try { if (!String(sh.getRange(1, 8).getValue() || '').trim()) { sh.getRange(1, 8).setValue(RC_HEAD[7]).setFontWeight('bold'); sh.setColumnWidth(8, 320); } } catch (x) {} return sh; }
  if (!sh) sh = ss.insertSheet(tabName_(RC_TAB));
  sh.clear(); sh.getRange(1, 1, 1, RC_HEAD.length).setValues([RC_HEAD]).setFontWeight('bold');
  const rows = RC_SEED.map(function (r) { return r.concat(['Yes', '']); });
  sh.getRange(2, 1, rows.length + 200, 5).setNumberFormat('@');   // 1/2 stays 1/2 (not a date)
  sh.getRange(2, 1, rows.length, 7).setValues(rows);
  try { sh.setFrozenRows(1); sh.getRange(2, 6, rows.length + 60, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build()); [180, 220, 140, 120, 360, 70, 300, 320].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); }); tabLook_(sh); hideStatus_(ss, null); } catch (x) {}
  return sh;
}
const rcNum_ = function (v) { return /^\s*\d+([.\/-]\d+)*\s*$/.test(String(v || '')); };
// Reads the tab (cached 5 minutes; an edit of the tab refreshes it right away). Status: warnings only, nothing is changed in your rows.
function rcRead_(fresh) {
  const c = CacheService.getScriptCache();
  if (!fresh) { try { const h = c.get('rc2'); if (h) return JSON.parse(h); } catch (x) {} }
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = rcSetup_(ss), n = sh.getLastRow() - 1;
  const v = n > 0 ? sh.getRange(2, 1, n, 8).getDisplayValues() : [], by = {}, order = [], st = [];
  v.forEach(function (r) {
    const name = String(r[0]).trim(), ing = String(r[1]).trim(), me = String(r[2]).trim(), am = String(r[3]).trim(), no = String(r[4]).trim();
    if (!name && !ing && !me && !am && !no) { st.push(['']); return; }
    if (!name) { st.push(['⚠ Recipe is empty (not shown)']); return; }
    if (!pcfgActive_(r[5])) { st.push(['Hidden (Active = No)']); return; }
    const k = name.toUpperCase(); if (!by[k]) { by[k] = { name: name, note: '', prep: [] }; order.push(k); }
    const R = by[k];
    const pm = /(?:\/d\/|[?&]id=|^)([-\w]{25,})/.exec(String(r[7] || '').trim()); if (pm && !R.photo) R.photo = pm[1];
    if (!ing && !me && !am) { if (!R.prep.length && !R.note && no) R.note = no; else if (no) R.prep.push(no); st.push(['✓ OK (written line)']); return; }
    const line = [ing, me, am]; while (line.length > 1 && line[line.length - 1] === '') line.pop(); if (no) { while (line.length < 3) line.push(''); line.push(no); }
    R.prep.push(line);
    st.push([rcNum_(me) && am && !rcNum_(am) ? '⚠ Measurement and Amount look swapped (Measurement is a number). Shown as typed.' : '✓ OK']);
  });
  try { stWrite_(sh, 2, 7, st); } catch (x) {}
  const list = order.map(function (k) { return by[k]; }).filter(function (r) { return r.prep.length || r.note; });
  const out = { ok: true, v: Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(list))).slice(0, 12), recipes: list };
  try { c.put('rc2', JSON.stringify(out), 300); } catch (x) {}
  return out;
}
/* ===================== COLLAB weekly schedule (W) =====================
   One horizontal tab, left to right: BUSINESS HOURS · EMPLOYEE WEEKLY TOTAL HOURS · COLLABORATORS (name, position, shift, active, o/a)
   · WEEKLY SCHEDULE (MON..SUN start / end) · WEEKLY COVERAGE (green blocks, 6 AM to midnight, 15 minutes per column, all 7 days).
   Titles on row 5, people from row 6. The app reads the columns by their titles. */
const TC2 = { hr: 5, r0: 6, n: 40, bh: 1, brk: 5, tot: 10, col: 13, sch: 20, ch: 43, chW: 72 };   // X layout. The W layout (col 9 / sch 16 / ch 32, no Role columns) is upgraded by the menu item
const TC2_BRK = [[0, 3, 0, 0], [3, 5, 1, 0], [5, 6, 1, 1], [6, 9, 2, 1]];   // shift hours from, up to (not included), paid 10-min breaks, unpaid 30-min meals (the paper table)
const TC2_BRKH = ['Shift length from (hrs)', 'Up to (hrs, not included)', 'Paid 10-min breaks', 'Unpaid 30-min meals'];
const TC2_DAYS = ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'], TC2_D3 = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];
const TC2_COLS = ['Collaborator Name', 'Position', 'Shift', 'Active', 'Spanish role ending (o / a)', 'Status (filled in automatically)'];
const TC2_NAVY = '#1F2A44', TC2_BLUE = '#2F4A7A', TC2_LIGHT = '#EEF2F8', TC2_GREEN = '#6AA84F';
function tc2Times_() { const o = []; for (let m = 360; m <= 1440; m += 15) o.push(tc2Fmt_(m)); return o; }
function tc2Fmt_(m) { m = m % 1440; const h = Math.floor(m / 60), mm = m % 60, h12 = h % 12 === 0 ? 12 : h % 12; return h12 + ':' + ('0' + mm).slice(-2) + ' ' + (h < 12 ? 'AM' : 'PM'); }
function tc2Col_(c) { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
// "6:00 AM", "6:00:00 PM", "18:15" or a time cell -> minutes after midnight (null = blank or not a time)
function tc2Min_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? null : v.getHours() * 60 + v.getMinutes();
  if (typeof v === 'number') return v >= 0 && v < 1.0001 ? Math.round(v * 1440) % 1440 : null;
  const m = /^\s*(\d{1,2})(?::(\d{2}))?(?::\d{2})?\s*([AaPp])?\.?\s*[Mm]?\.?\s*$/.exec(String(v));
  if (!m) return null; let h = +m[1]; const mm = +(m[2] || 0); if (mm > 59) return null;
  if (m[3]) { if (h < 1 || h > 12) return null; h = h % 12 + (/p/i.test(m[3]) ? 12 : 0); } else if (h > 23) return null;
  return h * 60 + mm;
}
function tcIsSched_(sh) { try { if (!sh) return false; const h = sh.getRange(TC2.hr, 9, 1, 5).getValues()[0].map(function (x) { return String(x || '').trim().toLowerCase(); }); return h[0] === 'collaborator name' || h[4] === 'collaborator name'; } catch (e) { return false; } }
function tcIsNew_(sh) { try { return !!sh && String(sh.getRange(TC2.hr, TC2.col).getValue() || '').trim().toLowerCase() === 'collaborator name'; } catch (e) { return false; } }
// The people on the tab: [{row, name, role, shift, act, rend, wk}] (wk only on the schedule layout: 7 x [start, end] minutes or null)
function tcRead_(sh) {
  if (!tcIsSched_(sh)) {
    const n = sh.getLastRow() - 1, v = n > 0 ? sh.getRange(2, 1, n, 6).getValues() : [];
    return { sched: false, r0: 2, stCol: 5, rows: v.map(function (r, i) { return { row: i + 2, name: r[0], role: r[1], shift: r[2], act: r[3], rend: r[5] }; }) };
  }
  const W = Math.max(sh.getLastColumn(), TC2.sch + 23), head = sh.getRange(TC2.hr, 1, 1, W).getValues()[0].map(function (x) { return String(x || '').trim().toLowerCase(); });
  const find = function (t, from) { for (let c = (from || 0); c < head.length; c++) if (head[c] === t.toLowerCase()) return c; return -1; };
  const cN = find('Collaborator Name'), cP = find('Position', cN), cS = find('Shift', cN), cA = find('Active', cN), cE = find('Spanish role ending (o / a)', cN), cT = find('Status (filled in automatically)', cN);
  const days = TC2_D3.map(function (d) { return [find(d + ' Start'), find(d + ' End'), find(d + ' Role')]; });
  const n = Math.max(0, sh.getLastRow() - TC2.hr), v = n ? sh.getRange(TC2.r0, 1, n, W).getValues() : [];
  const rows = []; let last = 0;
  v.forEach(function (r, i) { if (String(r[cN] || '').trim()) last = i + 1; });
  v.slice(0, last).forEach(function (r, i) {
    const bad = [];
    const wk = days.map(function (d, k) {
      if (d[0] < 0 || d[1] < 0) return null;
      const a = tc2Min_(r[d[0]]), b = tc2Min_(r[d[1]]), ea = r[d[0]] !== '' && r[d[0]] !== null, eb = r[d[1]] !== '' && r[d[1]] !== null;
      if (!ea && !eb) return null;
      if (a === null || b === null) { bad.push(TC2_D3[k] + ': pick both a start and an end'); return null; }
      let e2 = b <= a ? b + 1440 : b; if (e2 - a > 16 * 60) { bad.push(TC2_D3[k] + ': end time looks wrong'); return null; }
      const rp = d[2] >= 0 ? twPosOf_(r[d[2]]) : '';   // X: the position picked for that day (blank = the Main position)
      return rp ? [a, e2, rp] : [a, e2];
    });
    const tx = function (v) { if (v === '' || v === null || v === undefined) return ''; const m = tc2Min_(v); return m === null ? String(v) : tc2Fmt_(m); };
    const raw = days.map(function (d) { return [d[0] >= 0 ? tx(r[d[0]]) : '', d[1] >= 0 ? tx(r[d[1]]) : '', d[2] >= 0 ? String(r[d[2]] || '') : '']; });
    rows.push({ row: TC2.r0 + i, name: r[cN], role: cP >= 0 ? r[cP] : '', shift: cS >= 0 ? r[cS] : '', act: cA >= 0 ? r[cA] : '', rend: cE >= 0 ? r[cE] : '', wk: wk.some(Boolean) ? wk : null, raw: raw, bad: bad });
  });
  return { sched: true, r0: TC2.r0, stCol: cT >= 0 ? cT + 1 : 0, rows: rows };
}
function tcStWrite_(sh, T, st) { if (!T.stCol || !st.length) return; try { stWrite_(sh, T.r0, T.stCol, st); } catch (x) {} }
function tcBuildMenu() { const r = tcBuild_(SpreadsheetApp.getActiveSpreadsheet()); try { SpreadsheetApp.getActiveSpreadsheet().toast(r, 'Birriaz', 10); } catch (e) {} return r; }
// Turns the COLLAB tab into the weekly schedule. A copy of the old tab is kept first (COLLAB BACKUP date). Names, roles, shift, active and o/a move over.
function tcBuild_(ss) {
  let sh = tabFind_(ss, TC_TAB);
  if (sh && tcIsNew_(sh)) return 'COLLAB already has the newest layout. Nothing changed.';
  const upgrade = !!(sh && tcIsSched_(sh));
  const today = logNow_().date, old = sh ? tcRead_(sh).rows.filter(function (r) { return String(r.name || '').trim(); }) : [];
  if (sh) { let nm = tabName_(TC_TAB) + ' BACKUP ' + today, k = 2; while (ss.getSheetByName(nm)) nm = tabName_(TC_TAB) + ' BACKUP ' + today + ' (' + (k++) + ')'; sh.copyTo(ss).setName(nm); }
  else sh = ss.insertSheet(tabName_(TC_TAB));
  // Business hours the person already typed (the W layout kept them in A..C rows 6..12)
  let bhv = null; if (upgrade) { try { bhv = sh.getRange(TC2.r0, 2, 7, 2).getValues(); } catch (e) {} }
  const N = TC2.n, R0 = TC2.r0, lastC = TC2.ch + 7 * TC2.chW;
  try { sh.clearConditionalFormatRules(); } catch (e) {}
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); } catch (e) {}
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (e) {}
  sh.clear();
  try { sh.getBandings().forEach(function (b) { b.remove(); }); } catch (e) {}
  try { sh.showColumns(1, sh.getMaxColumns()); } catch (e) {}
  if (sh.getMaxColumns() < lastC) sh.insertColumnsAfter(sh.getMaxColumns(), lastC - sh.getMaxColumns());
  if (sh.getMaxRows() < R0 + N + 2) sh.insertRowsAfter(sh.getMaxRows(), R0 + N + 2 - sh.getMaxRows());
  const sch = TC2.sch, ch = TC2.ch, times = tc2Times_();
  // titles
  sh.getRange(1, 1, 1, lastC).setBackground(TC2_NAVY);
  sh.getRange(1, 1).setValue('COLLAB · TEAM & WEEKLY SCHEDULE').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(16);
  sh.getRange(2, 1).setValue('Add a person in PRIMARY COLLABORATORS (column ' + tc2Col_(TC2.col) + '). Pick start and end times from the dropdowns (6:00 AM to 12:00 AM) and, if the person works a different position that day, pick the Role. Totals and the green chart fill in by themselves. A day with no Role uses the Main position (Cashier, Cook or Dishes), which decides who gets ANY CASHIER / ANY COOK / ANY DISHES tasks. The BREAK AND MEAL TABLE can be edited.').setFontColor('#333333').setFontSize(10);
  const sec = [[TC2.bh, 3, 'BUSINESS HOURS'], [TC2.brk, 4, 'BREAK AND MEAL TABLE'], [TC2.tot, 2, 'EMPLOYEE WEEKLY TOTAL HOURS'], [TC2.col, 6, 'PRIMARY COLLABORATORS, POSITION AND SHIFT'], [sch, 22, 'WEEKLY SCHEDULE'], [ch, 1 + 7 * TC2.chW, 'WEEKLY COVERAGE (6 AM to midnight)']];
  sec.forEach(function (x) { const r = sh.getRange(3, x[0], 1, x[1]); try { r.merge(); } catch (e) {} r.setValue(x[2]).setBackground(TC2_NAVY).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center'); });
  // headers (row 5)
  const H = {}; H[TC2.bh] = ['Day', 'Open', 'Close']; H[TC2.brk] = TC2_BRKH; H[TC2.tot] = ['Employee Name', 'Total Hrs This Week']; H[TC2.col] = TC2_COLS;
  H[sch] = ['Name'].concat([].concat.apply([], TC2_D3.map(function (d) { return [d + ' Start', d + ' End', d + ' Role']; })));
  Object.keys(H).forEach(function (c) { sh.getRange(TC2.hr, +c, 1, H[c].length).setValues([H[c]]).setBackground(TC2_BLUE).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true); sh.getRange(4, +c, 1, H[c].length).setBackground(TC2_BLUE); });
  sh.getRange(TC2.hr, ch).setValue('Name').setBackground(TC2_BLUE).setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange(4, ch).setBackground(TC2_BLUE);
  for (let k = 0; k < 7; k++) {
    const c0 = ch + 1 + k * TC2.chW, dr = sh.getRange(4, c0, 1, TC2.chW); try { dr.merge(); } catch (e) {}
    dr.setValue(TC2_DAYS[k]).setBackground(k % 2 ? TC2_BLUE : TC2_NAVY).setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    for (let h = 0; h < 18; h++) { const hr = sh.getRange(TC2.hr, c0 + h * 4, 1, 4); try { hr.merge(); } catch (e) {} const hh = 6 + h; hr.setValue((hh % 12 === 0 ? 12 : hh % 12) + (hh < 12 ? 'a' : 'p')).setFontSize(8).setBackground(TC2_BLUE).setFontColor('#FFFFFF').setHorizontalAlignment('center'); }
    try { sh.getRange(4, c0, N + 2, 1).setBorder(null, true, null, null, null, null, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK); } catch (e) {}
  }
  // business hours
  sh.getRange(R0, TC2.bh, 7, 1).setValues(TC2_DAYS.map(function (d) { return [d]; })).setFontWeight('bold');
  if (bhv) { try { sh.getRange(R0, TC2.bh + 1, 7, 2).setNumberFormat('@').setValues(bhv); } catch (e) {} }
  // break and meal table (the paper table; the numbers can be edited)
  sh.getRange(R0, TC2.brk, TC2_BRK.length, 4).setValues(TC2_BRK).setHorizontalAlignment('center');
  sh.getRange(R0 + TC2_BRK.length + 1, TC2.brk).setValue('9 hours or more: no breaks are shown. Add a row above if there is a rule for it.').setFontColor('#666666').setFontSize(9);
  // people (moved from the old tab)
  const ppl = old.slice(0, N).map(function (r) { const a = String(r.act === true ? 'Yes' : r.act === false ? 'No' : r.act || '').trim(); return [String(r.name).trim().toUpperCase(), String(r.role || ''), String(r.shift || ''), /^(no|n|false|0|off|inactive|inactivo)$/i.test(a) ? 'No' : 'Yes', String(r.rend || '').trim().toLowerCase(), '']; });
  if (ppl.length) sh.getRange(R0, TC2.col, ppl.length, 6).setValues(ppl);
  // times and roles the person already had (upgrade): Name column stays a formula
  const L = function (c) { return tc2Col_(c); };
  const dvT = SpreadsheetApp.newDataValidation().requireValueInList(times, true).setAllowInvalid(false).setHelpText('Pick a time: 6:00 AM to 12:00 AM, every 15 minutes.').build();
  sh.getRange(R0, TC2.bh + 1, 7, 2).setNumberFormat('@').setDataValidation(dvT);
  sh.getRange(R0, sch + 1, N, 21).setNumberFormat('@');
  for (let k = 0; k < 7; k++) { sh.getRange(R0, sch + 1 + 3 * k, N, 2).setDataValidation(dvT); sh.getRange(R0, sch + 3 + 3 * k, N, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Cashier', 'Cook', 'Dishes'], true).setAllowInvalid(true).setHelpText('Leave blank to use the Main position.').build()); }
  if (upgrade && old.length) {
    const vals = old.slice(0, N).map(function (r) { const row = []; for (let k = 0; k < 7; k++) { const x = (r.raw && r.raw[k]) || ['', '', '']; row.push(x[0] === null || x[0] === undefined ? '' : String(x[0]), x[1] === null || x[1] === undefined ? '' : String(x[1]), x[2] === null || x[2] === undefined ? '' : String(x[2])); } return row; });
    sh.getRange(R0, sch + 1, vals.length, 21).setValues(vals);
  }
  // formulas: totals, names on the schedule and the chart
  const f = [];
  for (let i = 0; i < N; i++) {
    const r = R0 + i, I = '$' + L(TC2.col) + r;
    const terms = TC2_D3.map(function (d, k) { const a = '$' + L(sch + 1 + 3 * k) + r, b = '$' + L(sch + 2 + 3 * k) + r; return 'IF(OR(' + a + '="",' + b + '=""),0,MOD(TIMEVALUE(TO_TEXT(' + b + '))-TIMEVALUE(TO_TEXT(' + a + ')),1)*24)'; });
    f.push(['=IF(' + I + '="","",' + I + ')', '=IF(' + I + '="","",ROUND(' + terms.join('+') + ',2))']);
  }
  sh.getRange(R0, TC2.tot, N, 2).setFormulas(f);
  sh.getRange(R0, sch, N, 1).setFormulas(f.map(function (x) { return [x[0]]; }));
  sh.getRange(R0, ch, N, 1).setFormulas(f.map(function (x) { return [x[0]]; }));
  sh.getRange(R0, TC2.tot + 1, N, 1).setNumberFormat('0.00');
  // dropdowns
  sh.getRange(R0, TC2.col + 1, N, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Cashier', 'Cook', 'Dishes'], true).setAllowInvalid(true).build());
  sh.getRange(R0, TC2.col + 2, N, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['AM', 'MID', 'PM'], true).setAllowInvalid(true).build());
  sh.getRange(R0, TC2.col + 3, N, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(true).build());
  sh.getRange(R0, TC2.col + 4, N, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['o', 'a'], true).setAllowInvalid(true).build());
  // look: light rows, formulas grey, people bold
  sh.getRange(R0, 1, N, ch).setBackground('#FFFFFF');
  for (let i = 0; i < N; i += 2) sh.getRange(R0 + i, 1, 1, ch).setBackground(TC2_LIGHT);
  [TC2.bh + 3, TC2.brk + 4, TC2.tot + 2, TC2.col + 6, sch + 22].forEach(function (c) { sh.getRange(3, c, N + 3, 1).setBackground('#FFFFFF'); });
  sh.getRange(R0, TC2.tot, N, 2).setFontColor('#555555');
  sh.getRange(R0, TC2.col, N, 1).setFontWeight('bold');
  for (let k = 0; k < 7; k++) sh.getRange(R0, sch + 3 + 3 * k, N, 1).setBackground('#FFF4DC');
  sh.getRange(R0, ch + 1, N, 7 * TC2.chW).setBackground('#FAFAFA');
  // green coverage blocks: one rule per day (the cell's time = 6 AM + 15 minutes per column)
  const rules = [];
  for (let k = 0; k < 7; k++) {
    const c0 = ch + 1 + k * TC2.chW, a = '$' + L(sch + 1 + 3 * k) + R0, b = '$' + L(sch + 2 + 3 * k) + R0, tm = '(0.25+(COLUMN()-' + c0 + ')/96)';
    const fx = '=AND(LEN(' + a + ')>0,LEN(' + b + ')>0,' + tm + '>=TIMEVALUE(TO_TEXT(' + a + ')),' + tm + '<IF(TIMEVALUE(TO_TEXT(' + b + '))<=TIMEVALUE(TO_TEXT(' + a + ')),1,TIMEVALUE(TO_TEXT(' + b + '))))';
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(fx).setBackground(TC2_GREEN).setRanges([sh.getRange(R0, c0, N, TC2.chW)]).build());
  }
  try { sh.setConditionalFormatRules(rules); } catch (e) {}
  // widths, freeze titles, hide Status
  const wd = [[1, 110], [2, 84], [3, 84], [4, 14], [5, 110], [6, 110], [7, 100], [8, 100], [9, 14], [10, 130], [11, 100], [12, 14], [13, 150], [14, 100], [15, 70], [16, 70], [17, 120], [18, 220], [19, 14], [20, 120], [42, 14], [43, 110]];
  wd.forEach(function (x) { try { sh.setColumnWidth(x[0], x[1]); } catch (e) {} });
  try { sh.setColumnWidths(sch + 1, 21, 78); sh.setColumnWidths(ch + 1, 7 * TC2.chW, 8); } catch (e) {}
  for (let k = 0; k < 7; k++) { try { sh.setColumnWidth(sch + 3 + 3 * k, 72); } catch (e) {} }
  try { sh.setRowHeight(TC2.hr, 42); sh.setFrozenRows(TC2.hr); sh.setFrozenColumns(0); } catch (e) {}
  try { sh.hideColumns(TC2.col + 5); } catch (e) {}
  try { sh.setTabColor('#1F2A44'); } catch (e) {}
  try { CacheService.getScriptCache().remove('sh1'); } catch (e) {}
  let msg = 'COLLAB ' + (upgrade ? 'updated to the newest layout (Role columns, break table)' : 'is now the weekly schedule') + ' (' + ppl.length + ' people moved over). The old tab was copied first.';
  try { const r = twSheetSync_(); if (r && r.ok) msg += ' The tablets were updated.'; } catch (e) {}
  return msg;
}

// ---- X: Today's shifts (Home tile). Reads the schedule for today and the break table; the tablet shows it read-only. ----
function brkRead_(sh) {
  const d = TC2_BRK.map(function (x) { return { from: x[0], to: x[1], b10: x[2], meal: x[3] }; });
  try {
    if (!tcIsNew_(sh)) return d;
    const v = sh.getRange(TC2.r0, TC2.brk, 12, 4).getValues(), out = [];
    v.forEach(function (r) { if (r[0] === '' || r[1] === '' || isNaN(+r[0]) || isNaN(+r[1])) return; out.push({ from: +r[0], to: +r[1], b10: +r[2] || 0, meal: +r[3] || 0 }); });
    return out.length ? out : d;
  } catch (e) { return d; }
}
function shiftsGet_() {
  const date = logNow_().date, c = CacheService.getScriptCache();
  try { const h = c.get('sh1'); if (h) { const o = JSON.parse(h); if (o.date === date) return o; } } catch (x) {}
  let ss; try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { return { ok: false, error: 'no spreadsheet' }; }
  const sh = tabFind_(ss, TC_TAB); if (!sh) return { ok: true, date: date, people: [] };
  const T = tcRead_(sh), tb = brkRead_(sh), idx = twDayIdx_(date), people = [];
  T.rows.forEach(function (r) {
    const n = String(r.name || '').trim().toUpperCase(); if (!n) return;
    const a = String(r.act === true ? 'yes' : r.act || '').trim();
    if (/^(no|n|false|0|off|inactive|inactivo)$/i.test(a)) return;
    const main = twPosOf_(r.role), shift = String(r.shift || '').trim().toUpperCase();
    if (r.wk) {
      const w = r.wk[idx]; if (!w) return;   // has hours this week but none today
      const hrs = (w[1] - w[0]) / 60, row = tb.filter(function (x) { return hrs >= x.from && hrs < x.to; })[0];
      people.push({ name: n, pos: w[2] || main, s: w[0], e: w[1], hrs: Math.round(hrs * 100) / 100, b10: row ? row.b10 : null, meal: row ? row.meal : null });
    } else {
      people.push({ name: n, pos: main, s: null, e: null, shift: shift });   // no times on the schedule at all: only the Shift (AM / PM) is known
    }
  });
  people.sort(function (p, q) { return (p.s === null ? 9999 : p.s) - (q.s === null ? 9999 : q.s) || p.name.localeCompare(q.name); });
  const out = { ok: true, date: date, people: people };
  try { c.put('sh1', JSON.stringify(out), 60); } catch (x) {}
  return out;
}

// RECIPES: one blank row between recipes (only rows are inserted; nothing is changed or deleted). A copy of the tab is kept first.
function rcSpaceOut() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = tabFind_(ss, RC_TAB);
  if (!sh || sh.getLastRow() < 3) { try { ss.toast('RECIPES has nothing to space out.', 'Birriaz', 6); } catch (e) {} return 0; }
  const today = logNow_().date; let nm = 'RECIPES BACKUP ' + today, k = 2; while (ss.getSheetByName(nm)) nm = 'RECIPES BACKUP ' + today + ' (' + (k++) + ')';
  sh.copyTo(ss).setName(nm);
  const n = sh.getLastRow() - 1, names = sh.getRange(2, 1, n, 5).getDisplayValues(), at = [];
  let prev = '';
  for (let i = 0; i < n; i++) {
    const r = names[i], blank = r.every(function (x) { return String(x).trim() === ''; }), nm2 = String(r[0]).trim().toUpperCase();
    if (blank) { prev = ''; continue; }
    if (nm2 && prev && nm2 !== prev) at.push(i + 2);   // a new recipe starts right under the previous one
    if (nm2) prev = nm2;
  }
  for (let j = at.length - 1; j >= 0; j--) sh.insertRowBefore(at[j]);
  try { rcBand_(ss, null); } catch (e) {}
  try { CacheService.getScriptCache().remove('rc2'); } catch (e) {}
  try { ss.toast(at.length ? at.length + ' blank rows added between recipes. Backup: ' + nm + '.' : 'Recipes were already spaced out. Backup: ' + nm + '.', 'Birriaz', 8); } catch (e) {}
  return at.length;
}
// Collaborators tab, column F: the Spanish ending of the role (Cocinero / Cocinera). Added once to existing tabs; no row is changed.
function tcRendCol_(sh) {
  if (tcIsSched_(sh)) return;
  try {
    if (String(sh.getRange(1, 6).getValue() || '').trim() === '') {
      sh.getRange(1, 6).setValue(TC_HEAD[5]).setFontWeight('bold');
      try { sh.setColumnWidth(6, 190); } catch (x) {}
    }
    const n = Math.max(1, sh.getMaxRows() - 1);
    sh.getRange(2, 6, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['o', 'a'], true).setAllowInvalid(true).build());
  } catch (x) {}
}
// Training videos: every video file in the TRAINING VIDEOS folder (inside BIRRIAZ SOPs). The title is the file name before " - EN" / " - ES".
const TV_FOLDER_ID = '10fWqhUCLDiUSG58KOIfaKHHHda7uYxhK';
// RECIPES: every recipe block (rows that share the Recipe name, blank rows included) gets one light tint; the next recipe gets the other. Only colors change, no row moves.
function rcBand_(ss, out) {
  const sh = tabFind_(ss, RC_TAB); if (!sh || sh.getLastRow() < 2) return;
  const n = sh.getLastRow() - 1, names = sh.getRange(2, 1, n, 1).getDisplayValues(), tints = ['#FFF4DC', '#E8F1FB'], cols = [[1, 6], [8, 1]];
  let cur = '', k = -1; const bg = [];
  for (let i = 0; i < n; i++) { const nm = String(names[i][0]).trim().toUpperCase(); if (nm && nm !== cur) { cur = nm; k++; } bg.push([tints[(k < 0 ? 0 : k) % 2]]); }
  cols.forEach(function (c) { if (c[0] > sh.getMaxColumns()) return; const w = c[1], rows = bg.map(function (b) { const r = []; for (let j = 0; j < w; j++) r.push(b[0]); return r; }); sh.getRange(2, c[0], n, w).setBackgrounds(rows); });
  if (out) out.push('RECIPES colored by recipe');
}
function tvFolder_() {
  try { return DriveApp.getFolderById(TV_FOLDER_ID); } catch (e) {}
  try { const roots = zelFolders_(); for (let i = 0; i < roots.length; i++) { const it = DriveApp.getFolderById(roots[i]).getFolders(); while (it.hasNext()) { const f = it.next(); if (/^training videos?$/i.test(f.getName().trim())) return f; } } } catch (e) {}
  return null;
}
function tvRead_(fresh) {
  const c = CacheService.getScriptCache();
  if (!fresh) { try { const h = c.get('tv1'); if (h) return JSON.parse(h); } catch (x) {} }
  const by = {}, order = [], fo = tvFolder_();
  if (fo) {
    const it = fo.getFiles(); let n = 0;
    while (it.hasNext() && n++ < 300) {
      const f = it.next(); if (String(f.getMimeType() || '').indexOf('video/') !== 0) continue;
      let nm = f.getName().replace(/\.[A-Za-z0-9]{2,4}$/, '').trim(), lang = 'en';
      const m = /[\s_-]*[-–(\[]?\s*\b(EN|ES|ENG|ESP|English|Espa[ñn]ol|Spanish)\b\s*[)\]]?\s*$/i.exec(nm);
      if (m) { lang = /^(es|esp|espa|span)/i.test(m[1]) ? 'es' : 'en'; nm = nm.slice(0, m.index).trim(); }
      nm = nm.replace(/[\s_-]+$/, '');
      const k = nm.toLowerCase(); if (!k) continue;
      if (!by[k]) { by[k] = { title: nm }; order.push(k); }
      if (!by[k][lang]) by[k][lang] = f.getId();
    }
  }
  const list = order.map(function (k) { return by[k]; }).sort(function (a, b) { return a.title.localeCompare(b.title); });
  const out = { ok: true, found: !!fo, v: Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(list))).slice(0, 12), videos: list };
  try { c.put('tv1', JSON.stringify(out), 300); } catch (x) {}
  return out;
}
function tvGet_(v) { const r = tvRead_(false); return v && v === r.v ? { ok: true, v: r.v, same: true } : r; }
function rcNorm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim(); }
function rcTool_(name) {
  const list = rcRead_(false).recipes || [], n = rcNorm_(name);
  if (!n) return { recipes: list.map(function (r) { return r.name; }) };
  const hit = list.filter(function (r) { return rcNorm_(r.name) === n; })[0] || list.filter(function (r) { const k = rcNorm_(r.name); return k.indexOf(n) >= 0 || n.indexOf(k) >= 0; })[0];
  if (!hit) return { not_found: true, recipes: list.map(function (r) { return r.name; }) };
  const o = { recipe: hit.name, source: 'RECIPES tab', note: hit.note || '', lines: hit.prep.map(function (p) { return Array.isArray(p) ? { ingredient: p[0] || '', measurement: p[1] || '', amount: p[2] || '', notes: p[3] || '' } : { step: p }; }) };
  if (hit.photo) o.photo_ids = [hit.photo];
  return o;
}
// Photos the tablets may open: only the ones Zelyna knows (menu Docs) or a RECIPES row links to. Sent as a 1000 px picture.
function zphOk_(id) {
  id = String(id || ''); if (!/^[-\w]{20,80}$/.test(id)) return false;
  try { const d = zkbIdxRead_(); if (d && d.docs && Object.keys(d.docs).some(function (k) { return (d.docs[k].ph || []).some(function (p) { return p.f === id; }); })) return true; } catch (e) {}
  try { if ((rcRead_(false).recipes || []).some(function (r) { return r.photo === id; })) return true; } catch (e) {}
  return false;
}
function zphGet_(id) {
  if (!zphOk_(id)) return { ok: false, error: 'Photo not found.' };
  try {
    const tok = ScriptApp.getOAuthToken(), m = JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '?fields=thumbnailLink,mimeType,size', { headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true }).getContentText() || '{}');
    if (m.thumbnailLink) { const r = UrlFetchApp.fetch(String(m.thumbnailLink).replace(/=s\d+$/, '=s1000'), { headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true }); if (r.getResponseCode() === 200) { const b = r.getBlob(); return { ok: true, type: b.getContentType() || 'image/jpeg', data: Utilities.base64Encode(b.getBytes()) }; } }
    const f = DriveApp.getFileById(id); if (f.getSize() > 4500000) return { ok: false, error: 'Photo too big.' };
    const b = f.getBlob(); return { ok: true, type: b.getContentType() || 'image/jpeg', data: Utilities.base64Encode(b.getBytes()) };
  } catch (e) { return { ok: false, error: 'Photo not found.' }; }
}
function rcGet_(v) { const r = rcRead_(false); return v && v === r.v ? { ok: true, v: r.v, same: true } : r; }
function mrAdd_(raw) {
  const text = String((raw && raw.text) || '').trim().slice(0, 600);
  if (!raw || !raw.id || !text) return { ok: false, error: 'Write a description first.' };
  const id = String(raw.id).replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
  const d = mrRead_();
  if (d.items.some(function (x) { return x.id === id; })) return { ok: true, dup: true };
  let photo = '';
  if (raw.photo) {
    try {
      const b = Utilities.newBlob(Utilities.base64Decode(String(raw.photo)), 'image/jpeg', 'mr-' + id + '.jpg');
      photo = logFolder_().createFile(b).getId();
    } catch (x) { return { ok: false, error: 'The photo could not be saved: ' + x }; }
  }
  d.items.unshift({ id: id, text: text, emp: String(raw.emp || '').slice(0, 40), em: !!raw.em, photo: photo, by: String(raw.by || '').slice(0, 40), cts: String(raw.cts || logStamp_()).slice(0, 30), done: false });
  d.items = d.items.slice(0, 400);
  logWriteJson_(MR_FILE, d);
  try {
    getTab_('Issues', ['Date', 'Issue', 'Assigned to', 'Emergency', 'Photo', 'Done']).appendRow([new Date(), text, String(raw.emp || ''), raw.em ? 'YES' : '', photo ? 'yes' : '', '']);
  } catch (x) {}
  return { ok: true };
}
function mrDone_(raw) {
  const d = mrRead_(), it = d.items.filter(function (x) { return x.id === String((raw && raw.id) || ''); })[0];
  if (!it) return { ok: false, error: 'Issue not found.' };
  it.done = !!raw.done; it.doneAt = it.done ? logStamp_() : ''; it.doneBy = it.done ? String(raw.by || '').slice(0, 40) : '';
  logWriteJson_(MR_FILE, d);
  return { ok: true };
}
// Delete button on a completed assigned task: it leaves the app. The record (and its photo in Drive) is kept; the ISSUES tab is not changed.
function mrDel_(raw) {
  const d = mrRead_(), it = d.items.filter(function (x) { return x.id === String((raw && raw.id) || ''); })[0];
  if (!it) return { ok: true, gone: true };
  if (!it.done) return { ok: false, error: 'Only completed tasks can be deleted.' };
  it.removed = true; it.removedAt = logStamp_();
  logWriteJson_(MR_FILE, d);
  return { ok: true };
}
function mrPhoto_(id) {
  const it = mrRead_().items.filter(function (x) { return x.id === String(id || ''); })[0];
  if (!it || !it.photo) return { ok: false, error: 'No photo.' };
  try { return { ok: true, data: Utilities.base64Encode(DriveApp.getFileById(it.photo).getBlob().getBytes()) }; }
  catch (x) { return { ok: false, error: 'Photo not found.' }; }
}

function mrSetupTab_(ss, out) {
  const sh = getTab_('Issues', ['Date', 'Issue', 'Assigned to', 'Emergency', 'Photo', 'Done']);
  try { if (!sh.getBandings().length) tabLook_(sh); } catch (e) {}
  out.push('ISSUES tab ready.');
}


/* ===================== FIFO LABELS tab: the label names for the tablets =====================
   Left section (A-C) = REGULAR LABELS (product name, shelf life, show).  Right section (E-F) = BLACK STICKERS (white letters on black).
   Add, change or hide a row and the tablets get it on the next sync (edit trigger or the 30-minute check).
   If the REGULAR LABELS section is empty the tablets use the TODAY'S TASK - PREP LIST tab as before. */
const LBL_TAB = 'FIFO LABELS', LBL_FIRST = 3;
const LBL_BLACK_SEED = ['USE FIRST', 'ASADA', 'CHICKEN', 'AL PASTOR', 'CHEESE', 'CHURROS', 'WHITE RICE'];
// Columns of the FIFO LABELS tab (one management tab, three sections side by side):
//   REGULAR LABELS  A name | B category | C shelf life (days) | D show | E status (automatic)
//   BLACK STICKERS  F text | G show
//   CATEGORY DEFAULTS  I category | J default shelf life (days)   (used only when the product has no shelf life of its own)
const LC = { name: 1, cat: 2, sl: 3, show: 4, st: 5, bk: 6, bkShow: 7, dcat: 9, ddays: 10 };
const LBL_HEAD_REG = ['Label name', 'Category', 'Shelf life (days)', 'Show (Yes / No)', 'Status (filled in automatically)'];
function lblKey_(s) { return String(s || '').toUpperCase().replace(/\s*[–—-]\s*/g, ' – ').replace(/\s+/g, ' ').trim(); }
function lblName_(i) { return String(i.p ? i.p + (i.s ? ' – ' + i.s : '') : i.n || '').trim().toUpperCase().replace(/\s+/g, ' ').slice(0, 40); }
function lblTitles_(sh) {
  try { sh.getRange(1, 1, 1, 12).breakApart(); } catch (e) {}
  sh.getRange(1, LC.name, 1, 5).merge().setValue('REGULAR LABELS (name, used-by date, prepared date + time, who)').setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#45818E');
  sh.getRange(1, LC.bk, 1, 2).merge().setValue('BLACK STICKERS (white letters on black + prepared date and time)').setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#000000');
  sh.getRange(1, LC.dcat, 1, 2).merge().setValue('CATEGORY DEFAULTS (used only when a product has no shelf life of its own)').setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#38761D');
  sh.getRange(2, LC.name, 1, 5).setValues([LBL_HEAD_REG]).setFontWeight('bold').setBackground('#D0E0E3');
  sh.getRange(2, LC.bk, 1, 2).setValues([['Sticker text', 'Show (Yes / No)']]).setFontWeight('bold').setBackground('#D9D9D9');
  sh.getRange(2, LC.dcat, 1, 2).setValues([['Category', 'Default shelf life (days)']]).setFontWeight('bold').setBackground('#D9EAD3');
  try {
    sh.setFrozenRows(2); [[1, 300], [2, 170], [3, 130], [4, 120], [5, 360], [6, 240], [7, 120], [8, 30], [9, 200], [10, 170]].forEach(function (w) { sh.setColumnWidth(w[0], w[1]); });
    const yn = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build();
    sh.getRange(LBL_FIRST, LC.sl, 400, 1).clearDataValidations();
    sh.getRange(LBL_FIRST, LC.show, 400, 1).setDataValidation(yn); sh.getRange(LBL_FIRST, LC.bkShow, 400, 1).setDataValidation(yn);
    sh.setTabColor('#000000');
  } catch (e) {}
}
// One time: the tab made by the earlier version (no Category column) gets a Category column after the name.
// Nothing is retyped: the column is inserted, so every value, note and format moves with its column. A backup copy is saved first.
function lblMigrate_(sh, notes) {
  if (String(sh.getRange(2, 1).getValue() || '') !== 'Label name' || String(sh.getRange(2, 2).getValue() || '') === 'Category') return false;
  try { const n = Math.max(2, sh.getLastRow()); logWriteJson_('fifo-labels-backup-' + logNow_().date + '.json', { at: logStamp_(), values: sh.getRange(1, 1, n, 6).getValues() }); } catch (e) {}
  try { sh.getRange(1, 1, 1, 12).breakApart(); } catch (e) {}
  sh.insertColumnAfter(1);
  lblTitles_(sh);
  if (notes) notes.push('FIFO tab: Category column and CATEGORY DEFAULTS section added (all existing rows kept; backup saved).');
  return true;
}
function lblSetup_(ss, cfg, out) {
  let sh = tabFind_(ss, LBL_TAB);
  if (sh && sh.getLastRow() >= LBL_FIRST) { lblMigrate_(sh, out); out.push('FIFO tab already has rows, left as is.'); return; }
  if (!sh) sh = ss.insertSheet(tabName_(LBL_TAB));
  if (!cfg || !cfg.items) { try { cfg = pcfgFromSheet_(ss).cfg; } catch (e) {} }
  const reg = [], seen = {};
  ((cfg && cfg.items) || []).forEach(function (i) {
    if (!i || i.f === false) return;
    const n = lblName_(i), k = lblKey_(n);
    if (!n || seen[k]) return; seen[k] = 1; reg.push([n, '', i.sl || '', 'Yes']);
  });
  lblSeen_(reg.map(function (r) { return r[0]; }));
  const blk = LBL_BLACK_SEED.map(function (n) { return [n, 'Yes']; });
  lblTitles_(sh); try { tabLook_(sh); } catch (e) {}
  if (reg.length) sh.getRange(LBL_FIRST, LC.name, reg.length, 4).setValues(reg);
  sh.getRange(LBL_FIRST, LC.bk, blk.length, 2).setValues(blk);
  out.push('FIFO tab created (' + reg.length + ' regular labels, ' + blk.length + ' black stickers).');
}
// Names already offered once (so a row the manager deletes is not added back).
function lblSeen_(add) {
  const P = PropertiesService.getScriptProperties(); let a = [];
  try { a = JSON.parse(P.getProperty('LBL_SEEN') || '[]') || []; } catch (e) {}
  if (add && add.length) { add.forEach(function (n) { const k = lblKey_(n); if (a.map(lblKey_).indexOf(k) < 0) a.push(n); }); P.setProperty('LBL_SEEN', JSON.stringify(a.slice(-400))); }
  return a.map(lblKey_);
}
// A new PREP LIST item with "FIFO label" ticked is added to the REGULAR LABELS section once (never twice, also with other spacing or dashes).
function lblAddNew_(ss, cfg, notes) {
  const sh = tabFind_(ss, LBL_TAB); if (!sh || !cfg || !cfg.items) return;
  lblMigrate_(sh, notes);
  const last = Math.max(sh.getLastRow(), LBL_FIRST - 1), col = last >= LBL_FIRST ? sh.getRange(LBL_FIRST, LC.name, last - LBL_FIRST + 1, 1).getValues().map(function (r) { return lblKey_(r[0]); }) : [];
  const have = {}; col.forEach(function (n) { if (n) have[n] = 1; });
  const seen = lblSeen_(), add = [];
  cfg.items.forEach(function (i) { if (i.f === false) return; const n = lblName_(i), k = lblKey_(n); if (n && !have[k] && seen.indexOf(k) < 0) { have[k] = 1; add.push([n, '', i.sl || '', 'Yes']); } });
  if (!add.length) return;
  let row = LBL_FIRST; for (let k = col.length - 1; k >= 0; k--) { if (col[k]) { row = LBL_FIRST + k + 1; break; } }
  sh.getRange(row, LC.name, add.length, 4).setValues(add);
  lblSeen_(add.map(function (r) { return r[0]; }));
  notes.push('FIFO tab: added ' + add.length + ' new label(s) from the PREP LIST (category and shelf life to review).');
}
// Reads the three sections. Shelf life of a product: its own value (column C) > its category default (I-J) > none (the tablet warns and uses the standard 6 days).
// The FIFO tab is independent of the PREP LIST: names and shelf life are only read here. Column E shows why a row is not shown.
function lblRead_(ss, cfg) {
  const sh = tabFind_(ss, LBL_TAB); if (!sh) return null;
  if (cfg) lblMigrate_(sh, null);
  const n = sh.getLastRow() - LBL_FIRST + 1; if (n < 1) return { reg: [], blk: [], missing: 0 };
  const off = {}, cur = {}, seen = cfg ? lblSeen_() : [];
  ((cfg && cfg.items) || []).forEach(function (i) { const k = lblKey_(lblName_(i)); cur[k] = 1; if (i.f === false) off[k] = 1; });
  const v = sh.getRange(LBL_FIRST, 1, Math.min(n, 400), LC.ddays).getValues(), reg = [], blk = [], sr = {}, sb = {}, st = [], dflt = {};
  v.forEach(function (r) { const c = String(r[LC.dcat - 1] || '').trim().toUpperCase(), d = Math.round(+r[LC.ddays - 1]); if (c && d >= 1 && d <= 60 && !dflt[c]) dflt[c] = d; });
  let missing = 0;
  v.forEach(function (r) {
    const a = String(r[LC.name - 1] || '').trim().toUpperCase().replace(/\s+/g, ' ').slice(0, 40), k = lblKey_(a), cat = String(r[LC.cat - 1] || '').trim().toUpperCase().slice(0, 40);
    const raw = String(r[LC.sl - 1] === null || r[LC.sl - 1] === undefined ? '' : r[LC.sl - 1]).trim();
    let why = '';
    if (!a) why = '';
    else if (sr[k]) why = 'Listed twice (only the first row is used)';
    else if (!pcfgActive_(r[LC.show - 1])) why = 'Hidden (Show = No)';
    else {
      let d = Math.round(+raw), src = 'item';
      const bad = raw !== '' && !(d >= 1 && d <= 60);
      if (raw === '' || bad) { d = cat && dflt[cat] ? dflt[cat] : 0; src = d ? 'category' : 'none'; }
      sr[k] = 1; reg.push({ n: a, sl: d, src: src, cat: cat });
      if (src === 'none') missing++;
      why = (bad ? '⚠ Shelf life must be a whole number of days from 1 to 60. ' : '') +
        (src === 'item' ? '✓ On the tablets · ' + d + ' day' + (d > 1 ? 's' : '') + ' (this product)'
          : src === 'category' ? '✓ On the tablets · ' + d + ' day' + (d > 1 ? 's' : '') + ' (category ' + cat + ')'
          : '⚠ On the tablets · NO APPROVED SHELF LIFE: prints with the standard 6 days until you set column C' + (cat ? ' or a default for ' + cat : ' or a category default'));
    }
    st.push([why]);
    const b = String(r[LC.bk - 1] || '').trim().toUpperCase().replace(/\s+/g, ' ').slice(0, 30);
    if (b && !sb[b] && pcfgActive_(r[LC.bkShow - 1])) { sb[b] = 1; blk.push(b); }
  });
  try { if (cfg && st.length) sh.getRange(LBL_FIRST, LC.st, st.length, 1).setValues(st); } catch (e) {}
  const cats = []; reg.forEach(function (x) { if (x.cat && cats.indexOf(x.cat) < 0) cats.push(x.cat); });
  return { reg: reg, blk: blk, cats: cats, missing: missing };
}


/* ===================== VOICE MESSAGES for each collaborator (Collaborator task list -> profile) =====================
   The audio is saved as a file in the data folder. A message stays until someone taps Delete (never deleted automatically). */
const VM_FILE = 'voice-msgs.json';
function vmRead_() { const d = logReadJson_(VM_FILE); return (d && Array.isArray(d.items)) ? d : { items: [] }; }
function vmList_() { return { ok: true, items: vmRead_().items.map(function (x) { return { id: x.id, emp: x.emp, by: x.by || '', cts: x.cts, dur: x.dur || 0, heard: !!x.heard }; }) }; }
function vmAdd_(raw) {
  raw = raw || {};
  const id = String(raw.id || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40), emp = String(raw.emp || '').trim().toUpperCase().slice(0, 40), data = String(raw.data || '');
  if (!id || !emp || !data) return { ok: false, error: 'Nothing to save.' };
  if (data.length > 6000000) return { ok: false, error: 'The message is too long. Keep it under 2 minutes.' };
  const d = vmRead_(); if (d.items.some(function (x) { return x.id === id; })) return { ok: true, dup: true };
  const mime = /^audio\/[a-z0-9.+-]+/i.test(String(raw.mime || '')) ? String(raw.mime).split(';')[0] : 'audio/webm';
  let file = '';
  try { file = logFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(data), mime, 'voice-' + emp.replace(/\W+/g, '_') + '-' + id + '.' + (mime.split('/')[1] || 'webm'))).getId(); }
  catch (x) { return { ok: false, error: 'The message could not be saved: ' + x }; }
  d.items.unshift({ id: id, emp: emp, by: String(raw.by || '').slice(0, 40), cts: String(raw.cts || logStamp_()).slice(0, 30), dur: Math.max(0, Math.min(600, Math.round(+raw.dur || 0))), mime: mime, file: file, heard: false });
  d.items = d.items.slice(0, 300); logWriteJson_(VM_FILE, d);
  return { ok: true };
}
function vmAudio_(id) {
  const it = vmRead_().items.filter(function (x) { return x.id === String(id || ''); })[0];
  if (!it) return { ok: false, error: 'Message not found.' };
  try { return { ok: true, mime: it.mime || 'audio/webm', data: Utilities.base64Encode(DriveApp.getFileById(it.file).getBlob().getBytes()) }; }
  catch (x) { return { ok: false, error: 'Audio file not found.' }; }
}
function vmHeard_(raw) { const d = vmRead_(), it = d.items.filter(function (x) { return x.id === String((raw && raw.id) || ''); })[0]; if (!it) return { ok: false }; it.heard = true; logWriteJson_(VM_FILE, d); return { ok: true }; }
function vmDel_(raw) {
  const d = vmRead_(), id = String((raw && raw.id) || ''), it = d.items.filter(function (x) { return x.id === id; })[0];
  if (!it) return { ok: true, gone: true };
  try { DriveApp.getFileById(it.file).setTrashed(true); } catch (x) {}
  d.items = d.items.filter(function (x) { return x.id !== id; }); logWriteJson_(VM_FILE, d);
  return { ok: true };
}

/* ===================== WEEKLY TRAINING (1-2 minute team huddle written by Zelyna) =====================
   One lesson per week (Monday to Sunday). The TRAINING PLAN tab shows this week and the next 4 weeks:
   type a Topic and/or Notes (something that happened, what to stress) for any week; leave Topic empty and Zelyna picks
   the next SOP document or ZELYNA Q&A category in turn. Every lesson also looks at what happened the week before
   (failed temperatures, maintenance issues, questions staff asked) and connects to it when it fits.
   Each collaborator plays it in their profile and signs. When everyone has signed (or Sunday 9:15 PM) the report is emailed. */
const TN_FILE = 'training.json', TN_TAB = 'TRAINING PLAN', TN_AHEAD = 4;
const TN_HEAD = ['Week of (Monday)', 'Topic (empty = Zelyna picks)', 'Notes for Zelyna (what happened, what to stress)', 'Source (filled in automatically)', 'Lesson (filled in automatically)', 'Completions (filled in automatically)'];
function tnRead_() { const d = logReadJson_(TN_FILE); const o = (d && d.weeks) ? d : { weeks: {}, used: (d && d.used) || [] }; o.used = o.used || []; return o; }
function tnWeek_() { return logMonday_(logNow_().date); }
function tnNames_() { try { return (twConfig_().employees || []).filter(function (e) { return !e.off; }).map(function (e) { return e.name; }); } catch (e) { return []; } }
function tnMdY_(ds) { const p = String(ds).split('-'); return p[1] + '/' + p[2] + '/' + p[0]; }
function tnSources_() {
  const out = [];
  try {
    zelDocList_(80).forEach(function (d) { if (d.type !== 'Text') out.push({ key: 'doc:' + d.id, name: d.name, kind: 'sop', id: d.id }); });
  } catch (e) {}
  const by = {};
  zelFaqRows_().forEach(function (r) { const q = String(r[0] || '').trim(), a = String(r[1] || '').trim(); if (q && a && zelOn_(r[2])) { const c = String(r[3] || '').trim() || 'Other'; (by[c] = by[c] || []).push('Q: ' + q + '\nA: ' + a); } });
  Object.keys(by).forEach(function (c) { out.push({ key: 'qa:' + c, name: c, kind: 'qa', text: by[c].join('\n') }); });
  return out;
}
// ---- TRAINING PLAN tab
function tnPlanSheet_() { return tabFind_(SpreadsheetApp.getActiveSpreadsheet(), TN_TAB); }
function tnPlanEnsure_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let sh = tabFind_(ss, TN_TAB), made = false;
  if (!sh) {
    sh = ss.insertSheet(tabName_(TN_TAB)); made = true;
    sh.getRange(1, 1, 1, TN_HEAD.length).setValues([TN_HEAD]).setFontWeight('bold').setBackground('#674EA7').setFontColor('#FFFFFF');
    try { sh.setFrozenRows(1); sh.setColumnWidth(1, 130); sh.setColumnWidth(2, 280); sh.setColumnWidth(3, 380); sh.setColumnWidth(4, 260); sh.setColumnWidth(5, 300); sh.setColumnWidth(6, 220); sh.getRange(2, 2, 300, 2).setWrap(true); } catch (e) {}
    try { tabLook_(sh); } catch (e) {}
  }
  const have = {}, n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, 1).getValues().forEach(function (r) { const k = tsDate_(r[0]); if (k) have[logMonday_(k)] = 1; });
  const cur = tnWeek_(), add = [];
  for (let k = 0; k <= TN_AHEAD; k++) { const w = logAddDays_(cur, 7 * k); if (!have[w]) add.push([tnMdY_(w), '', '', '', '', '']); }
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, TN_HEAD.length).setValues(add);
  return { sheet: sh, made: made, added: add.length };
}
function tnPlanRead_() {
  const sh = tnPlanSheet_(), out = {}; if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r, i) {
    const k = tsDate_(r[0]); if (!k) return; const w = logMonday_(k);
    out[w] = { row: i + 2, topic: String(r[1] || '').trim().slice(0, 150), notes: String(r[2] || '').trim().slice(0, 1200) };
  });
  return out;
}
function tnPk_(p) { return (p && (p.topic || p.notes)) ? (p.topic + '|' + p.notes) : ''; }
// Which source each upcoming week will use (weeks with a typed Topic use it; the others take the next source in turn).
function tnPredict_(d, plan, src) {
  const cur = tnWeek_(), used = (d.used || []).slice(), out = {};
  let left = src.filter(function (x) { return used.indexOf(x.key) < 0; });
  for (let k = 0; k <= TN_AHEAD; k++) {
    const w = logAddDays_(cur, 7 * k), W = d.weeks[w], p = plan[w] || {};
    if (p.topic) { out[w] = { kind: 'custom', name: p.topic }; continue; }
    if (W && W.srcKey) { const f = src.filter(function (x) { return x.key === W.srcKey; })[0]; if (f) { out[w] = f; continue; } }
    if (!left.length) left = src.slice();
    if (left.length) { out[w] = left.shift(); }
  }
  return out;
}
function tnSrcTxt_(x) { return !x ? 'Nothing to teach yet (add SOPs or Q&A answers)' : x.kind === 'custom' ? 'Your topic: ' + x.name : x.kind === 'sop' ? 'SOP "' + x.name + '"' : 'ZELYNA Q&A: ' + x.name; }
function tnPlanStatus_() {
  try {
    const r0 = tnPlanEnsure_(), sh = r0.sheet; if (sh.getLastRow() < 2) return;
    const d = tnRead_(), plan = tnPlanRead_(), pred = tnPredict_(d, plan, tnSources_()), names = tnNames_(), cur = tnWeek_();
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues(), out = v.map(function (r) {
      const k = tsDate_(r[0]); if (!k) return ['', '', ''];
      const w = logMonday_(k), W = d.weeks[w], done = W ? Object.keys(W.done || {}).length : 0;
      const srcT = W && W.en ? tnSrcTxt_({ kind: W.kind, name: W.src }) : (w >= cur ? tnSrcTxt_(pred[w]) : '');
      const les = W && W.en ? '"' + W.en.title + '"' + (W.pk !== tnPk_(plan[w]) && !done ? ' (will be rewritten with your new topic/notes)' : '') : (w < cur ? 'No lesson' : w === cur ? 'Written on the first visit today' : 'Written Monday ' + tnMdY_(w) + ' at 5:30 AM (or tap "Write it now" in the app)');
      const cpl = W && W.en && w <= cur ? done + ' of ' + names.length + ' signed' + (W.mailed ? ' · email sent' : '') : '';
      return [srcT, les, cpl];
    });
    sh.getRange(2, 4, out.length, 3).setValues(out);
  } catch (e) { Logger.log('tnPlanStatus_: ' + e); }
}
// What happened the week before (feeds the lesson so Zelyna can connect to it).
function tnLastWeek_(wk) {
  const from = logAddDays_(wk, -7), to = logAddDays_(wk, -1), L = [];
  try { const e = logReadWeek_(from, false).entries.filter(function (x) { return x.result === 'FAIL'; }); if (e.length) L.push(e.length + ' failed temperature reading' + (e.length > 1 ? 's' : '') + ': ' + e.slice(0, 4).map(function (x) { return [x.equip, x.food, x.temp ? x.temp + ' F' : ''].filter(String).join(' '); }).join('; ')); } catch (x) {}
  try { const m = mrRead_().items.filter(function (x) { const d = String(x.cts || '').slice(0, 10); return d >= from && d <= to; }); if (m.length) L.push(m.length + ' maintenance/assigned issue' + (m.length > 1 ? 's' : '') + ': ' + m.slice(0, 4).map(function (x) { return x.text; }).join('; ')); } catch (x) {}
  try { const q = zqTop_(from, to, 3, false); if (q.items.length) L.push('Staff asked Zelyna most about: ' + q.items.map(function (x) { return x.theme + ' (' + x.count + ')'; }).join('; ')); } catch (x) {}
  try { const u = zelFaqRows_().filter(function (r) { return /Needs an answer/.test(String(r[4] || '')); }); if (u.length) L.push(u.length + ' question' + (u.length > 1 ? 's' : '') + ' Zelyna could not answer, for example: ' + u.slice(0, 2).map(function (r) { return r[0]; }).join('; ')); } catch (x) {}
  return L;
}
function tnJson_(txt) { const a = txt.indexOf('{'), b = txt.lastIndexOf('}'); return JSON.parse(txt.slice(a, b + 1)); }
function tnMaterial_(src) {
  if (src.kind === 'sop') return ((zelDocRead_(src.id) || {}).text || '').slice(0, 9000);
  if (src.kind === 'qa') return String(src.text || '').slice(0, 6000);
  // your own topic: the best matching SOP + matching Q&A rows
  let m = '';
  try { const f = (zelDocSearch_(src.name.split(/\s+/).filter(function (w) { return w.length > 3; })[0] || src.name).documents || [])[0]; if (f) m += 'SOP "' + f.name + '":\n' + ((zelDocRead_(f.id) || {}).text || '').slice(0, 6000) + '\n'; } catch (e) {}
  const words = src.name.toLowerCase().split(/\W+/).filter(function (w) { return w.length > 3; });
  zelFaqRows_().forEach(function (r) { const q = String(r[0] || ''), a = String(r[1] || ''); if (a && zelOn_(r[2]) && words.some(function (w) { return (q + ' ' + a).toLowerCase().indexOf(w) >= 0; })) m += 'Q: ' + q + '\nA: ' + a + '\n'; });
  return m.slice(0, 9000);
}
function tnWrite_(src, notes, last) {
  const material = tnMaterial_(src);
  if (zelKey_()) {
    try {
      const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': zelKey_(), 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: zelModel_(), max_tokens: 2500,
          system: 'You are Zelyna, the assistant of a Mexican restaurant (' + NOVI_LOC + '). You write a short weekly training huddle for the kitchen and cashier team, read aloud by the tablet. Simple words, short sentences, friendly and practical, no blame and no names. No markdown. Use the facts in the material; when the material does not cover something, keep to widely accepted restaurant and food safety practice and keep it simple.',
          messages: [{ role: 'user', content: (src.kind === 'custom' ? 'TOPIC CHOSEN BY THE OWNER: ' + src.name + '\n' : 'TOPIC SOURCE: ' + tnSrcTxt_(src) + '\n') +
            (notes ? 'OWNER NOTES (what happened / what to stress): ' + notes + '\n' : '') +
            (last.length ? 'LAST WEEK AT THE RESTAURANT:\n- ' + last.join('\n- ') + '\n' : '') +
            'MATERIAL:\n' + (material || '(no document matched)') +
            '\n\nWrite one lesson that takes about 1 minute to read aloud (120 to 150 words, never more than 150). If the owner notes or something from last week relates to the topic, open with one sentence about it. Then write exactly 3 multiple-choice questions about this lesson only, each with exactly 3 short choices and one right answer (the other two must be clearly wrong to someone who listened), plus one short line saying why the answer is right. The Spanish questions are the same questions in the same order with the same right answer. Reply with ONLY this JSON: {"title_en":"","title_es":"","script_en":"","script_es":"","points_en":["","",""],"points_es":["","",""],"quiz_en":[{"q":"","choices":["","",""],"answer":0,"why":""},{"q":"","choices":["","",""],"answer":0,"why":""},{"q":"","choices":["","",""],"answer":0,"why":""}],"quiz_es":[{"q":"","choices":["","",""],"answer":0,"why":""},{"q":"","choices":["","",""],"answer":0,"why":""},{"q":"","choices":["","",""],"answer":0,"why":""}],"link_en":"one short line saying which part of last week this connects to, or empty"}' }] }) });
      const j = JSON.parse(res.getContentText()), txt = (j.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
      const o = tnJson_(txt);
      const qz = tnQuiz_(o.quiz_en, o.quiz_es);
      const L = function (k) { return { title: String(o['title_' + k] || src.name).slice(0, 120), script: String(o['script_' + k] || '').slice(0, 2500), points: (o['points_' + k] || []).slice(0, 4).map(function (x) { return String(x).slice(0, 200); }), q: '', a: '', quiz: qz[k] }; };
      const r = { en: L('en'), es: L('es'), link: String(o.link_en || '').slice(0, 200) };
      if (r.en.script) return r;
    } catch (e) { Logger.log('Training AI failed: ' + e); }
  }
  const lines = (material || (src.name + '. ' + (notes || ''))).split('\n').map(function (x) { return x.trim(); }).filter(String);
  const ans = src.kind === 'qa' ? lines.filter(function (x) { return /^A:/.test(x); }).map(function (x) { return x.slice(2).trim(); }) : lines.slice(0, 6);
  const L0 = { title: src.name, script: ((notes ? notes + ' ' : '') + (src.kind === 'qa' ? lines.join(' ').replace(/\bQ: /g, ' Question: ').replace(/\bA: /g, ' Answer: ') : lines.slice(0, 14).join(' '))).slice(0, 1400), points: ans.slice(0, 3).map(function (x) { return x.slice(0, 200); }), q: '', a: '' };
  return { en: L0, es: L0, plain: true, link: '' };
}
// Makes (or re-makes) the lesson of a week. A lesson is rewritten when its Topic/Notes changed and nobody has signed yet.
function tnEnsure_(wk, force) {
  let d = tnRead_(), plan = tnPlanRead_(), p = plan[wk] || { topic: '', notes: '' }, W = d.weeks[wk];
  const signed = W && Object.keys(W.done || {}).length;
  if (W && W.en && (signed || (!force && W.pk === tnPk_(p)))) return d;
  const lock = LockService.getScriptLock(); try { lock.waitLock(30000); } catch (e) { return d; }
  try {
    d = tnRead_(); W = d.weeks[wk];
    if (W && W.en && Object.keys(W.done || {}).length) return d;
    const srcAll = tnSources_(), pred = tnPredict_(d, plan, srcAll);
    let src = p.topic ? { kind: 'custom', name: p.topic, key: 'custom:' + p.topic } : (W && W.srcKey ? srcAll.filter(function (x) { return x.key === W.srcKey; })[0] : null) || pred[wk];
    if (!src) { const left = srcAll.filter(function (x) { return d.used.indexOf(x.key) < 0; }); src = left[0] || srcAll[0]; }
    if (!src) return d;
    const last = tnLastWeek_(wk), lesson = tnWrite_(src, p.notes, last);
    d.weeks[wk] = { src: src.name, kind: src.kind, srcKey: src.key, pk: tnPk_(p), en: lesson.en, es: lesson.es, plain: !!lesson.plain, link: lesson.link || '', last: last, made: logStamp_(), done: (W && W.done) || {}, mailed: (W && W.mailed) || '' };
    if (src.kind !== 'custom' && d.used.indexOf(src.key) < 0) d.used = d.used.concat([src.key]).slice(-200);
    const keys = Object.keys(d.weeks).sort(); while (keys.length > 30) delete d.weeks[keys.shift()];
    logWriteJson_(TN_FILE, d);
    return d;
  } finally { try { lock.releaseLock(); } catch (e) {} }
}
function tnPub_(W) { return { src: W.src, kind: W.kind, en: W.en, es: W.es, plain: !!W.plain, link: W.link || '', done: Object.keys(W.done || {}).reduce(function (o, k) { o[k] = W.done[k].at; return o; }, {}), mailed: W.mailed || '' }; }
function tnGet_() {
  const wk = tnWeek_(), d = tnEnsure_(wk), W = d.weeks[wk];
  if (!W) return { ok: true, week: wk, none: true, names: tnNames_() };
  return Object.assign({ ok: true, week: wk, names: tnNames_() }, tnPub_(W));
}
function tnPlanGet_() {
  tnPlanEnsure_();
  const d = tnRead_(), plan = tnPlanRead_(), pred = tnPredict_(d, plan, tnSources_()), cur = tnWeek_(), weeks = [];
  for (let k = -1; k <= TN_AHEAD; k++) {
    const w = logAddDays_(cur, 7 * k), W = d.weeks[w], p = plan[w] || { topic: '', notes: '' };
    if (k < 0 && !W) continue;
    weeks.push(Object.assign({ week: w, topic: p.topic, notes: p.notes, source: W && W.en ? tnSrcTxt_({ kind: W.kind, name: W.src }) : tnSrcTxt_(pred[w]), stale: !!(W && W.en && W.pk !== tnPk_(p) && !Object.keys(W.done || {}).length) }, W && W.en ? { lesson: tnPub_(W) } : {}));
  }
  return { ok: true, cur: cur, names: tnNames_(), weeks: weeks, nextLink: tnLastWeek_(logAddDays_(cur, 7)) };
}
function tnPlanSet_(raw) {
  raw = raw || {}; if (String(raw.pin || '') !== TW_PIN) return { ok: false, error: 'wrong manager code' };
  const k0 = tsDate_(raw.week), w = k0 ? logMonday_(k0) : ''; if (!w || w < tnWeek_()) return { ok: false, error: 'Pick this week or a future week.' };
  tnPlanEnsure_(); const sh = tnPlanSheet_(), p = tnPlanRead_()[w]; if (!p) return { ok: false, error: 'Week not found.' };
  sh.getRange(p.row, 2, 1, 2).setValues([[String(raw.topic || '').trim().slice(0, 150), String(raw.notes || '').trim().slice(0, 1200)]]);
  tnPlanStatus_();
  return { ok: true };
}
// 3 multiple-choice questions (3 choices each, one right). Checked, then the choices are put in a mixed order
// (same order in English and Spanish) so the right answer is not always the first one. Anything wrong = no quiz for that lesson.
function tnQuiz_(en, es) {
  const ok = function (a) { return Array.isArray(a) && a.length === 3 && a.every(function (x) { return x && String(x.q || '').trim() && Array.isArray(x.choices) && x.choices.length === 3 && x.choices.every(function (c) { return String(c || '').trim(); }) && [0, 1, 2].indexOf(Number(x.answer)) >= 0; }); };
  if (!ok(en)) return { en: [], es: [] };
  if (!ok(es) || es.some(function (x, i) { return Number(x.answer) !== Number(en[i].answer); })) es = en;
  const out = { en: [], es: [] };
  en.forEach(function (x, i) {
    const order = [0, 1, 2].sort(function () { return Math.random() - 0.5; });
    [['en', x], ['es', es[i]]].forEach(function (p) {
      const y = p[1];
      out[p[0]].push({ q: String(y.q).trim().slice(0, 200), choices: order.map(function (k) { return String(y.choices[k]).trim().slice(0, 120); }), answer: order.indexOf(Number(y.answer)), why: String(y.why || '').trim().slice(0, 200) });
    });
  });
  return out;
}
function tnMake_(raw) {
  raw = raw || {}; if (String(raw.pin || '') !== TW_PIN) return { ok: false, error: 'wrong manager code' };
  const k0 = tsDate_(raw.week), w = k0 ? logMonday_(k0) : ''; if (!w || w < tnWeek_()) return { ok: false, error: 'Pick this week or a future week.' };
  const d = tnEnsure_(w, true), W = d.weeks[w]; tnPlanStatus_();
  return W && W.en ? { ok: true, lesson: tnPub_(W) } : { ok: false, error: 'Nothing to teach yet. Add SOPs to the BIRRIAZ SOPs folder or answers to the ZELYNA Q&A tab.' };
}
function tnSign_(raw) {
  raw = raw || {};
  const wk = tnWeek_(), name = String(raw.name || '').trim().toUpperCase(), names = tnNames_();
  if (names.indexOf(name) < 0) return { ok: false, error: 'Unknown collaborator.' };
  const d = tnEnsure_(wk), W = d.weeks[wk]; if (!W) return { ok: false, error: 'There is no training this week.' };
  W.done = W.done || {};
  if (W.done[name]) return { ok: true, dup: true };
  let sig = '';
  if (raw.sig) { try { sig = logFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(String(raw.sig)), 'image/png', 'training-' + wk + '-' + name.replace(/\W+/g, '_') + '.png')).getId(); } catch (x) {} }
  const nq = ((W.en && W.en.quiz) || []).length, sc = /^\d\/\d$/.test(String(raw.score || '')) && +String(raw.score).slice(0, 1) <= nq && +String(raw.score).slice(2) === nq ? String(raw.score) : '';
  W.done[name] = { at: String(raw.at || logStamp_()).slice(0, 30), sig: sig, score: sc };
  logWriteJson_(TN_FILE, d);
  try { const sh = getTab_('TRAINING LOG', ['Week of', 'Topic', 'Collaborator', 'Signed at', 'Signature file', 'Score (right on the first try)']); if (!String(sh.getRange(1, 6).getValue() || '').trim()) sh.getRange(1, 6).setValue('Score (right on the first try)').setFontWeight('bold');
    sh.appendRow([tnMdY_(wk), W.en.title, name, W.done[name].at, sig ? 'https://drive.google.com/file/d/' + sig : '', sc || (nq ? '' : 'no quiz')]); } catch (x) {}
  const all = names.every(function (n) { return W.done[n]; });
  if (all && !W.mailed) { try { tnMail_(wk, d); } catch (x) { Logger.log('Training email failed: ' + x); } }
  try { tnPlanStatus_(); } catch (x) {}
  return { ok: true, all: all };
}
function tnMail_(wk, d) {
  d = d || tnRead_(); const W = d.weeks[wk]; if (!W || W.mailed) return false;
  const names = tnNames_(), done = names.filter(function (n) { return W.done && W.done[n]; }), miss = names.filter(function (n) { return !(W.done && W.done[n]); });
  const tm = function (s) { s = String(s || ''); const m = /(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})/.exec(s); if (!m) return s; const h = +m[4]; return LOG_MON[+m[2] - 1] + ' ' + (+m[3]) + ', ' + ((h % 12) || 12) + ':' + m[5] + (h < 12 ? ' AM' : ' PM'); };
  const L = ['Hi, this is Zelyna, your restaurant operations assistant from ' + NOVI_LOC + '.', '', 'Here is the weekly training report for the week of ' + logFmt_(wk) + '.', '',
    "THIS WEEK'S TOPIC", '- ' + W.en.title, '- Source: ' + tnSrcTxt_({ kind: W.kind, name: W.src })];
  if (W.link) L.push('- Connected to last week: ' + W.link);
  L.push('', 'KEY POINTS'); (W.en.points || []).forEach(function (p) { L.push('- ' + p); });
  if ((W.en.quiz || []).length) { L.push('', 'QUESTIONS'); W.en.quiz.forEach(function (x, i) { L.push((i + 1) + '. ' + x.q, '   Answer: ' + x.choices[x.answer] + (x.why ? ' (' + x.why + ')' : '')); }); }
  else if (W.en.q) L.push('', 'QUICK QUESTION', '- ' + W.en.q, '- Answer: ' + W.en.a);
  L.push('', 'COMPLETIONS', '- ' + done.length + ' of ' + names.length + ' collaborators (' + (names.length ? Math.round(done.length * 100 / names.length) : 0) + '%)' + (miss.length ? '' : ' - everyone finished'));
  done.forEach(function (n) { L.push('  ✓ ' + n + ' - signed ' + tm(W.done[n].at) + (W.done[n].score ? ' - score ' + W.done[n].score : '')); });
  if (miss.length) { L.push('', 'NOT COMPLETED'); miss.forEach(function (n) { L.push('  ✗ ' + n); }); }
  try { const p = tnPlanRead_()[logAddDays_(wk, 7)], pr = tnPredict_(d, tnPlanRead_(), tnSources_())[logAddDays_(wk, 7)]; L.push('', 'NEXT WEEK', '- ' + (p && p.topic ? 'Your topic: ' + p.topic : tnSrcTxt_(pr)) + ' (change it on the TRAINING tab)'); } catch (x) {}
  L.push('', 'The signatures are attached.', '', '— Zelyna', 'Your Restaurant Operations Assistant', 'Powered by Zelyna · Digital Pilot OPS', NOVI_LOC);
  const att = [];
  done.forEach(function (n) { const f = W.done[n].sig; if (f) { try { att.push(DriveApp.getFileById(f).getBlob().setName('signature-' + n.replace(/\W+/g, '_') + '.png')); } catch (x) {} } });
  MailApp.sendEmail({ to: REPORT_TO, subject: 'Zelyna | Weekly Training | ' + NOVI_LOC + ' | Week of ' + logFmt_(wk) + ' | ' + done.length + '/' + names.length + ' completed', body: L.join('\n'), attachments: att, name: NOVI_FROM });
  W.mailed = logStamp_(); logWriteJson_(TN_FILE, d);
  return true;
}
function tnDow_() { const p = logNow_().date.split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay(); }
function tnMorning() {
  try { tnPlanEnsure_(); tnEnsure_(tnWeek_()); tnPlanStatus_(); } catch (e) { Logger.log('tnMorning: ' + e); }
  try { if (tnDow_() === 1) zqWeekly_(logAddDays_(tnWeek_(), -7)); } catch (e) { Logger.log('questions email: ' + e); }
}
function tnEvening() { try { if (tnDow_() !== 0) return; const wk = tnWeek_(), d = tnRead_(); if (d.weeks[wk] && !d.weeks[wk].mailed) tnMail_(wk, d); } catch (e) { Logger.log('tnEvening: ' + e); } }
function tnTriggers_(out) {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (['tnMorning', 'tnEvening'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tnMorning').timeBased().everyDays(1).atHour(5).nearMinute(30).inTimezone(LOG_TZ).create();
  ScriptApp.newTrigger('tnEvening').timeBased().everyDays(1).atHour(21).nearMinute(15).inTimezone(LOG_TZ).create();
  try { tnPlanEnsure_(); tnPlanStatus_(); } catch (e) {}
  if (out) out.push('Weekly training installed (TRAINING tab; lesson Monday 5:30 AM; report when everyone signs or Sunday 9:15 PM; most-asked questions email Monday 5:30 AM).');
}

/* ===================== QUESTIONS staff ask Zelyna (log + weekly "most asked" email) ===================== */
const ZQ_FILE = 'zelyna-questions.json';
function zqRead_() { const d = logReadJson_(ZQ_FILE); return (d && Array.isArray(d.items)) ? d : { items: [] }; }
function zqLog_(q, answered, lang, model) {
  q = String(q || '').trim().slice(0, 300);
  if (q.length < 6 || /^(yes|yeah|ok|okay|no|si|sí|claro|dale|cancel|thanks|thank you|gracias)\b/i.test(q)) return;
  const lock = LockService.getScriptLock(); try { lock.waitLock(5000); } catch (e) { return; }
  try {
    const d = zqRead_(); d.items.push({ at: logStamp_(), q: q, a: !!answered, lang: lang || 'en', m: model || '' }); d.items = d.items.slice(-3000); logWriteJson_(ZQ_FILE, d);
    try { const sh = getTab_('ZELYNA QUESTIONS', ['Date', 'Question', 'Answered', 'Language', 'Model']); if (!String(sh.getRange(1, 5).getValue() || '').trim()) sh.getRange(1, 5).setValue('Model').setFontWeight('bold');
      sh.appendRow([logStamp_(), q, answered ? 'Yes' : 'No - saved on the ZELYNA Q&A tab', lang === 'es' ? 'Spanish' : 'English', model ? (/haiku/i.test(model) ? 'Fast (' + model + ')' : model) : '']); } catch (x) {}
  } finally { try { lock.releaseLock(); } catch (e) {} }
}
// Groups the questions of a date range into themes (AI when the key is set, otherwise same wording).
function zqTop_(from, to, max, useAi) {
  const list = zqRead_().items.filter(function (x) { const d = String(x.at).slice(0, 10); return d >= from && d <= to; });
  if (!list.length) return { total: 0, answered: 0, items: [] };
  let groups = null;
  if (useAi !== false && zelKey_() && list.length > 1) {
    try {
      const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': zelKey_(), 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: zelModel_(), max_tokens: 1500, system: 'You group restaurant staff questions that ask the same thing (any language). Reply with ONLY JSON: [{"theme":"short English theme","q":[question numbers]}]. Every question number goes in exactly one group.',
          messages: [{ role: 'user', content: list.slice(-400).map(function (x, i) { return (i + 1) + '. ' + x.q; }).join('\n') }] }) });
      const j = JSON.parse(res.getContentText()), txt = (j.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
      const arr = JSON.parse(txt.slice(txt.indexOf('['), txt.lastIndexOf(']') + 1)), L = list.slice(-400);
      groups = arr.map(function (g) { const qs = (g.q || []).map(function (n) { return L[n - 1]; }).filter(Boolean); return { theme: String(g.theme || '').slice(0, 120), count: qs.length, example: qs[0] ? qs[0].q : '', open: qs.filter(function (x) { return !x.a; }).length }; }).filter(function (g) { return g.count; });
    } catch (e) { groups = null; }
  }
  if (!groups) {
    const by = {};
    list.forEach(function (x) { const k = zelFaqNorm_(x.q); (by[k] = by[k] || { theme: x.q, count: 0, example: x.q, open: 0 }).count++; if (!x.a) by[k].open++; });
    groups = Object.keys(by).map(function (k) { return by[k]; });
  }
  groups.sort(function (a, b) { return b.count - a.count; });
  return { total: list.length, answered: list.filter(function (x) { return x.a; }).length, items: groups.slice(0, max || 10) };
}
function zqWeekly_(wk) {
  const P = PropertiesService.getScriptProperties(); if (P.getProperty('ZQ_SENT') === wk) return false;
  const to = logAddDays_(wk, 6), r = zqTop_(wk, to, 10, true);
  const open = zelFaqRows_().filter(function (r0) { return /Needs an answer/.test(String(r0[4] || '')); });
  const L = ['Hi, this is Zelyna, your restaurant operations assistant from ' + NOVI_LOC + '.', '', 'Here are the questions the team asked me during the week of ' + logFmt_(wk) + '.', '',
    'SUMMARY', '- ' + r.total + ' question' + (r.total === 1 ? '' : 's') + ' asked · ' + r.answered + ' answered · ' + (r.total - r.answered) + ' I could not answer', ''];
  L.push('MOST ASKED');
  if (!r.items.length) L.push('- Nobody asked me anything this week.');
  r.items.forEach(function (g, i) { L.push((i + 1) + '. ' + g.theme + ' - ' + g.count + ' time' + (g.count > 1 ? 's' : '') + (g.example && g.example !== g.theme ? ' (for example: "' + g.example + '")' : '') + (g.open ? ' · not answered ' + g.open + 'x' : '')); });
  L.push('', 'WAITING FOR YOUR ANSWER (ZELYNA Q&A tab, column B)');
  if (!open.length) L.push('- None. Every question has an answer.');
  open.slice(0, 15).forEach(function (r0) { L.push('- ' + r0[0] + ' · ' + String(r0[4]).replace(/^❓\s*/, '')); });
  if (r.items.length) L.push('', 'IDEA FOR NEXT WEEK', '- Make a training about "' + r.items[0].theme + '" (type it as the Topic on the TRAINING tab).');
  L.push('', '— Zelyna', 'Your Restaurant Operations Assistant', 'Powered by Zelyna · Digital Pilot OPS', NOVI_LOC);
  MailApp.sendEmail({ to: REPORT_TO, subject: 'Zelyna | Most asked questions | ' + NOVI_LOC + ' | Week of ' + logFmt_(wk), body: L.join('\n'), name: NOVI_FROM });
  P.setProperty('ZQ_SENT', wk);
  return true;
}
