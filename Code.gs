/**
 * FBA Prep Sequencer v2 — Apps Script backend
 * Bound to the "FBA Prep Sequencer - SKU Library" spreadsheet.
 *
 * Setup (one time):
 *   1. Extensions → Apps Script, replace everything with this file.
 *   2. Project Settings → Script Properties → add:
 *        SHARED_KEY   = a long random string (the app sends it on every request)
 *        NOTIFY_EMAIL = where issue and close-out emails go
 *   3. Deploy → New deployment → Web app → Execute as: Me · Who has access: Anyone → Deploy.
 *      Paste the /exec URL and the SHARED_KEY into the app's Settings tab.
 *   4. After any edit here: Deploy → Manage deployments → edit → New version → Deploy.
 *
 * Sheets used (created automatically):
 *   State      — one JSON blob per key (library, shipment, issues, log, settings, training). The app's source of truth.
 *   Training photos go to a Drive folder named "FBA Prep Training" (created on first use). Run setupTraining() once from the editor to grant the Drive permission.
 *   SKU Library, Runs, Issues, Log — readable copies, rewritten on each save. Edit the app, not these.
 */

const STATE_KEYS = ['library', 'shipment', 'issues', 'log', 'settings', 'training'];

function doGet(e) {
  return json_({ ok: true, app: 'fba-prep-sequencer', version: 2 });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  const key = PropertiesService.getScriptProperties().getProperty('SHARED_KEY');
  if (!key || body.key !== key) return json_({ ok: false, error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (body.action) {
      case 'load':  return json_({ ok: true, state: loadState_() });
      case 'save':  saveState_(body.state || {}); return json_({ ok: true, state: loadState_() });
      case 'issue': appendIssue_(body.issue); emailIssue_(body.issue); return json_({ ok: true });
      case 'close': appendLog_(body.entry); emailClose_(body.entry, body.detail); return json_({ ok: true });
      case 'photo': return json_(savePhoto_(body));
      case 'photoDelete': { try { DriveApp.getFileById(body.id).setTrashed(true); } catch (err) {} return json_({ ok: true }); }
      case 'test':  { const to = notify_(); if (!to) return json_({ ok: false, error: 'NOTIFY_EMAIL script property is not set' });
                      try { MailApp.sendEmail(to, '[Prep] Test email', 'The FBA Prep Sequencer can send mail. Sent ' + new Date().toLocaleString()); return json_({ ok: true, to: to }); }
                      catch (err) { return json_({ ok: false, error: 'mail failed: ' + err.message + ' — open the script editor, run testEmail() once, and approve the permission prompt' }); } }
      default: return json_({ ok: false, error: 'unknown action' });
    }
  } finally { lock.releaseLock(); }
}

// ── state ──────────────────────────────────────────────────────────────────
function stateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('State');
  if (!sh) { sh = ss.insertSheet('State'); sh.getRange(1, 1, 1, 3).setValues([['key', 'updated', 'json']]); sh.setFrozenRows(1); }
  return sh;
}
function loadState_() {
  const sh = stateSheet_(); const rows = sh.getDataRange().getValues(); const raw = {};
  for (let i = 1; i < rows.length; i++) { const [k, updated, js] = rows[i]; if (k) raw[k] = { updated: Number(updated) || 0, js: js || '' }; }
  const out = {};
  for (const k of STATE_KEYS) {
    if (!raw[k]) continue; let js = raw[k].js; let i = 2; while (raw[k + '#' + i] && raw[k + '#' + i].js) { js += raw[k + '#' + i].js; i++; }
    try { out[k] = { updated: raw[k].updated, data: js ? JSON.parse(js) : null }; } catch (err) { out[k] = { updated: raw[k].updated, data: null }; }
  }
  return out;
}
function saveState_(state) {
  const sh = stateSheet_(); const rows = sh.getDataRange().getValues(); const idx = {};
  for (let i = 1; i < rows.length; i++) idx[rows[i][0]] = i + 1;
  const now = Date.now();
  for (const k of STATE_KEYS) {
    if (!(k in state)) continue;
    const js = JSON.stringify(state[k] === undefined ? null : state[k]);
    if (js.length > 49000) { writeChunked_(sh, k, js, now, idx); continue; }
    const row = idx[k] || sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 3).setValues([[k, now, js]]);
    idx[k] = row;
  }
  if ('library' in state) writeLibrarySheets_(state.library || []);
  if ('issues' in state) writeIssuesSheet_(state.issues || []);
  if ('log' in state) writeLogSheet_(state.log || []);
}
// A cell holds 50k chars; the library with runs can exceed that, so split across key, key#2, key#3 … and rejoin on load.
function writeChunked_(sh, k, js, now, idx) {
  const parts = js.match(/[\s\S]{1,45000}/g);
  parts.forEach((p, i) => { const kk = i === 0 ? k : k + '#' + (i + 1); const row = idx[kk] || sh.getLastRow() + 1; sh.getRange(row, 1, 1, 3).setValues([[kk, now, p]]); idx[kk] = row; });
  // clear stale extra chunks
  let i = parts.length + 1; while (idx[k + '#' + i]) { sh.getRange(idx[k + '#' + i], 3).setValue(''); i++; }
}
// ── readable sheets ────────────────────────────────────────────────────────
function rewrite_(name, header, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
  sh.clearContents(); sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold'); sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
}
function writeLibrarySheets_(lib) {
  const rows = lib.map(p => {
    const ok = (p.runs || []).filter(r => r.flag === 'ok'); const u = ok.reduce((s, r) => s + r.units, 0), m = ok.reduce((s, r) => s + r.mins, 0);
    const rate = u ? m / u * 60 : (p.setRate ? p.setRate * 60 : '');
    return [p.name, p.sku, (p.altSkus || []).join(', '), p.asin, p.fnsku || '', p.upc, p.l, p.w, p.h, p.wt, p.unitWt, p.attrs, rate ? Number(rate.toFixed(1)) : '', ok.length, u];
  });
  rewrite_('SKU Library', ['Product', 'SKU', 'Old SKUs', 'ASIN', 'FNSKU', 'Units/carton', 'L', 'W', 'H', 'lb', 'Unit lb', 'Attributes', 'Sec/unit', 'Runs', 'Units timed'], rows);
  const runs = []; lib.forEach(p => (p.runs || []).forEach(r => runs.push([new Date(r.ts), p.name, p.sku, r.mins, r.units, Number((r.mins / r.units * 60).toFixed(1)), r.flag])));
  runs.sort((a, b) => b[0] - a[0]);
  rewrite_('Runs', ['When', 'Product', 'SKU', 'Minutes', 'Units', 'Sec/unit', 'Flag'], runs);
}
function writeIssuesSheet_(issues) {
  rewrite_('Issues', ['When', 'Shipment', 'Type', 'Qty', 'Product', 'SKU', 'ASIN', 'Ship qty', 'Where', 'Note', 'By', 'Resolved'],
    issues.map(i => [new Date(i.ts), i.shipId, i.type, i.qty, i.name, i.sku, i.asin, i.units, i.dest, i.text, i.who, i.resolved ? 'yes' : '']));
}
function writeLogSheet_(log) {
  rewrite_('Log', ['Closed', 'Shipment', 'Destination', 'Units', 'Boxes', 'Pallets', 'Plan lb', 'Actual lb', 'Est. prep min', 'Timed min', 'Issues', 'Fit', 'Film at end'],
    log.map(e => [new Date(e.ts), e.shipId, e.dest, e.units, e.boxes, e.pallets, e.planWt, e.actWt, e.est, e.timed, e.issues, e.fit, e.endFilm]));
}
function appendIssue_(i) { /* readable sheet is rewritten on the next save; nothing else needed */ }
function appendLog_(e) { }

// ── email ──────────────────────────────────────────────────────────────────
function notify_() { return PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL'); }
function emailIssue_(i) {
  const to = notify_(); if (!to || !i) return;
  const subj = `[Prep] ${i.type}${i.qty ? ' ×' + i.qty : ''} — ${i.name}${i.shipId ? ' (' + i.shipId + ')' : ''}`;
  const body = [
    `Issue: ${i.type}${i.qty ? ' — ' + i.qty + ' affected' : ''}`,
    `Product: ${i.name}`, `SKU: ${i.sku}   ASIN: ${i.asin}`,
    `Shipment: ${i.shipId || '-'}   qty in shipment: ${i.units}`, i.dest ? `Goes to: ${i.dest}` : '',
    i.text ? `\n${i.text}` : '', `\nReported by ${i.who || 'unknown'} at ${new Date(i.ts).toLocaleString()}`
  ].filter(Boolean).join('\n');
  MailApp.sendEmail(to, subj, body);
}
function emailClose_(e, d) {
  const to = notify_(); if (!to || !e) return;
  const wo = e.wo || (d && d.wo) || '';
  const subj = `[Prep] Shipment closed — ${wo ? 'WO ' + wo + ' · ' : ''}${e.shipId || 'no id'} · ${e.pallets} pallets · ${e.units} units`;
  const lines = [
    `${wo ? 'WO ' + wo + ' · ' : ''}Shipment ${e.shipId || '-'}${e.dest ? ' → ' + e.dest : ''}`, `Closed ${new Date(e.ts).toLocaleString()}`,
    d && d.started ? `Job ${d.started} → ${d.finished || '?'}${e.jobMins ? ' · ' + fmtMin_(e.jobMins) + ' · ' + e.secPerBox + ' sec/box' : ''}` : '', '',
    `Units ${e.units} · Boxes ${e.boxes} · Pallets ${e.pallets}`,
    `Weight: planned ${e.planWt} lb · measured ${e.actWt ? e.actWt + ' lb' : '(not all pallets weighed)'}`,
    `Prep time: estimated ${fmtMin_(e.est)}${e.estAll ? '' : ' (partial)'} · timed ${fmtMin_(e.timed)}`,
    `Issues logged: ${e.issues}   Fit: ${e.fit}   Film loaded at end: ${e.endFilm ? e.endFilm + '"' : '-'}`, ''
  ];
  if (d && d.sessions && d.sessions.length) { lines.push('Job clock' + (d.days > 1 ? ' (' + d.days + ' days)' : '') + ':'); d.sessions.forEach(x => lines.push(`  ${x.start} → ${x.stop}   ${fmtMin_(x.mins)}`)); lines.push(''); }
  if (d && d.pallets) { lines.push('Pallets (planned → actual):'); d.pallets.forEach(p => lines.push(`  ${p.n}: ${p.planWt} lb / ${p.planH}"  →  ${p.actWt ? p.actWt + ' lb' : '— lb'} / ${p.actH ? p.actH + '"' : '—'}   · ${p.boxes} boxes · ${p.layers} layers`)); lines.push(''); }
  if (d && d.products) { lines.push('Products:'); d.products.forEach(p => lines.push(`  ${p.units} u / ${p.boxes} bx  ${p.name}`)); lines.push(''); }
  if (d && d.issues && d.issues.length) { lines.push('Issues:'); d.issues.forEach(i => lines.push(`  ${i.type}${i.qty ? ' ×' + i.qty : ''} — ${i.name}${i.text ? ': ' + i.text : ''}`)); }
  MailApp.sendEmail(to, subj, lines.join('\n'));
}
// ── training photos ────────────────────────────────────────────────────────
const TRAINING_FOLDER = 'FBA Prep Training';
function trainingFolder_() { const it = DriveApp.getFoldersByName(TRAINING_FOLDER); if (it.hasNext()) return it.next(); const f = DriveApp.createFolder(TRAINING_FOLDER); return f; }
function savePhoto_(b) {
  if (!b || !b.b64) return { ok: false, error: 'no image' };
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(b.b64), b.mime || 'image/jpeg', b.name || ('photo-' + Date.now() + '.jpg'));
    const file = trainingFolder_().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // so the app can show it on any device
    return { ok: true, id: file.getId(), url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600' };
  } catch (err) { return { ok: false, error: 'Drive: ' + err.message + ' — open the script editor, run setupTraining() once, and approve the permission prompt' }; }
}
/** Run this once from the editor (Run ▸ setupTraining) to grant the Drive permission and create the photo folder. */
function setupTraining() { const f = trainingFolder_(); Logger.log('Training photo folder ready: ' + f.getUrl()); }

/** Run this once from the editor (Run ▸ testEmail) to grant the mail permission and check NOTIFY_EMAIL. */
function testEmail() { const to = notify_(); if (!to) throw new Error('Set the NOTIFY_EMAIL script property first'); MailApp.sendEmail(to, '[Prep] Test email', 'Mail permission is working. ' + new Date().toLocaleString()); Logger.log('sent to ' + to); }

function fmtMin_(m) { if (!m) return '-'; const h = Math.floor(m / 60), mm = Math.round(m % 60); return h ? `${h}h ${mm}m` : `${mm}m`; }

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
