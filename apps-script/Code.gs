/**
 * Fit Check logs: paste into the Google Sheet (Extensions > Apps Script), run setup once,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * The app sends one JSON row per API call, run, feedback and button tap.
 * Team memory: checks also save extracted reviews to the Reviews tab and read them back by brand (doGet ?action=reviews&brand=...).
 * After changing this code: Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy. The URL stays the same.
 */
const LOGS = 'Logs', KPIS = 'KPIs', CHUNK = 45000;
const HEAD = ['Time','Type','Run ID','Session','Version','Step / action','Provider','Model','Status','ms','Attempt','Fit call','Feedback','Detail','Request','Response','More (long text continues)'];

const REVIEWS = 'Reviews';
const RHEAD = ['Saved','Brand','Product','Source','Rating','Date','Size bought','Height (in)','Body type','Fit','Review text','Key'];
const rkey_ = (brand, text) => String(brand).toLowerCase().trim() + '|' + String(text).toLowerCase().replace(/\s+/g, ' ').trim().slice(0, 80);
function reviewsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let sh = ss.getSheetByName(REVIEWS);
  if (!sh) { sh = ss.insertSheet(REVIEWS); sh.getRange(1, 1, 1, RHEAD.length).setValues([RHEAD]).setFontWeight('bold'); sh.setFrozenRows(1); sh.setColumnWidth(11, 420); }
  return sh;
}
function saveReviews_(d) {
  const sh = reviewsSheet_(), last = sh.getLastRow();
  const have = new Set(last > 1 ? sh.getRange(2, 12, last - 1, 1).getValues().map(r => r[0]) : []);
  const out = [];
  (d.rows || []).forEach(r => { if (!r || !r.text || !r.brand) return; const k = rkey_(r.brand, r.text); if (have.has(k)) return; have.add(k);
    out.push([new Date(), r.brand, r.product || '', r.source || '', r.rating === '' ? '' : r.rating, r.date || '', r.size_bought || '', r.height_in === '' ? '' : r.height_in, r.body_type || '', r.fit || '', String(r.text).slice(0, 2000), k]); });
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, RHEAD.length).setValues(out);
  return out.length;
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.type === 'memory') { const n = saveReviews_(d); return ContentService.createTextOutput('saved ' + n); }
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOGS) || setup();
    const parts = s => { s = s == null ? '' : String(s); const out = []; for (let i = 0; i < s.length; i += CHUNK) out.push(s.slice(i, i + CHUNK)); return out.length ? out : ['']; };
    const req = parts(d.request), res = parts(d.response), det = String(d.detail || '').slice(0, CHUNK);
    const extra = req.slice(1).map((x, i) => 'REQUEST part ' + (i + 2) + ': ' + x).concat(res.slice(1).map((x, i) => 'RESPONSE part ' + (i + 2) + ': ' + x));
    const row = [d.time ? new Date(d.time) : new Date(), d.type || '', String(d.run_id || ''), d.session || '', d.version || '', d.step || '', d.provider || '', d.model || '',
      String(d.status || ''), d.ms || '', d.attempt || '', d.fit_call || '', d.feedback || '', det, req[0], res[0]].concat(extra);
    sh.appendRow(row);
    return ContentService.createTextOutput('ok');
  } finally { lock.releaseLock(); }
}
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'reviews') {
    const brand = String(p.brand || '').toLowerCase().trim(), sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REVIEWS);
    const rows = [];
    if (brand && sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, RHEAD.length).getValues().forEach(r => {
        if (String(r[1]).toLowerCase().trim() !== brand || rows.length >= 1500) return;
        rows.push({brand: r[1], product: r[2], source: r[3], rating: r[4], date: r[5] instanceof Date ? Utilities.formatDate(r[5], 'UTC', 'yyyy-MM-dd') : r[5], size_bought: r[6], height_in: r[7], body_type: r[8], fit: r[9], text: r[10]});
      });
    }
    return ContentService.createTextOutput(JSON.stringify({rows})).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('Fit Check log endpoint is running.');
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(LOGS);
  if (!sh) { sh = ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(LOGS) : ss.insertSheet(LOGS); }
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold'); sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss'); sh.setColumnWidths(14, 4, 320);
  buildKpis_(ss); return sh;
}

function buildKpis_(ss) {
  let k = ss.getSheetByName(KPIS); if (k) k.clear(); else k = ss.insertSheet(KPIS);
  const L = 'Logs!';
  const runs = (st) => `COUNTIFS(${L}B:B,"run",${L}I:I,"${st}")`;
  const rows = [
    ['Fit Check KPIs', ''],
    ['Checks started', `=COUNTIFS(${L}B:B,"action",${L}F:F,"Analyze")`],
    ['Checks finished', `=${runs('ok')}`],
    ['Finished %', '=IFERROR(B3/B2,"")'],
    ['Avg seconds per finished check', `=IFERROR(AVERAGEIFS(${L}J:J,${L}B:B,"run",${L}I:I,"ok")/1000,"")`],
    ['Finished in under 60 s %', `=IFERROR(COUNTIFS(${L}B:B,"run",${L}I:I,"ok",${L}J:J,"<60000")/B3,"")`],
    ['Checks with errors', `=${runs('error')}`],
    ['Couldn\'t read the page', `=${runs('unreadable')}`],
    ['Not enough evidence', `=${runs('not_enough')}`],
    ['Cancelled', `=${runs('cancelled')}`],
    ['API calls', `=COUNTIFS(${L}B:B,"api")`],
    ['API retries', `=COUNTIFS(${L}B:B,"api",${L}K:K,">1")`],
    ['API calls that failed', `=COUNTIFS(${L}B:B,"api",${L}I:I,"<>200")`],
    ['', ''],
    ['Fit calls shown', ''],
    ['Good fit', `=COUNTIFS(${L}B:B,"run",${L}L:L,"Good fit")`],
    ['Check first', `=COUNTIFS(${L}B:B,"run",${L}L:L,"Check first")`],
    ['Poor fit', `=COUNTIFS(${L}B:B,"run",${L}L:L,"Poor fit")`],
    ['', ''],
    ['Feedback', ''],
    ['👍 Useful', `=COUNTIFS(${L}B:B,"feedback",${L}M:M,"up")`],
    ['😐 Meh', `=COUNTIFS(${L}B:B,"feedback",${L}M:M,"meh")`],
    ['👎 Not useful', `=COUNTIFS(${L}B:B,"feedback",${L}M:M,"down")`],
    ['Useful %', '=IFERROR(B21/SUM(B21:B23),"")'],
    ['', ''],
    ['Avg seconds by step (successful calls)', ''],
    ['Reader', `=IFERROR(AVERAGEIFS(${L}J:J,${L}B:B,"api",${L}F:F,"Reader",${L}I:I,"200")/1000,"")`],
    ['Web researcher', `=IFERROR(AVERAGEIFS(${L}J:J,${L}B:B,"api",${L}F:F,"Web researcher",${L}I:I,"200")/1000,"")`],
    ['Analyst', `=IFERROR(AVERAGEIFS(${L}J:J,${L}B:B,"api",${L}F:F,"Analyst",${L}I:I,"200")/1000,"")`],
    ['Decider', `=IFERROR(AVERAGEIFS(${L}J:J,${L}B:B,"api",${L}F:F,"Decider",${L}I:I,"200")/1000,"")`],
    ['Photo estimator', `=IFERROR(AVERAGEIFS(${L}J:J,${L}B:B,"api",${L}F:F,"Photo estimator",${L}I:I,"200")/1000,"")`]
  ];
  k.getRange(1, 1, rows.length, 2).setValues(rows.map(r => [r[0], ''])); 
  rows.forEach((r, i) => { if (r[1]) k.getRange(i + 2 - 1, 2).setFormula(r[1]); });
  ['B4','B6','B24'].forEach(a => k.getRange(a).setNumberFormat('0%'));
  ['B5','B27:B31'].forEach(a => k.getRange(a).setNumberFormat('0.0'));
  [1,15,20,26].forEach(r => k.getRange(r, 1).setFontWeight('bold'));
  k.getRange('D1').setValue('Checks per day').setFontWeight('bold');
  k.getRange('D2').setFormula(`=IFERROR(QUERY(${L}A:I,"select toDate(A), count(A) where B = 'run' and I = 'ok' group by toDate(A) label toDate(A) 'Day', count(A) 'Finished checks'",0),"No checks yet")`);
  k.setColumnWidth(1, 280);
}
