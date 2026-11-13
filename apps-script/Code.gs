/**
 * Fit Check logs: paste into the Google Sheet (Extensions > Apps Script), run setup once,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * The app sends one JSON row per API call, run, feedback and button tap.
 */
const LOGS = 'Logs', KPIS = 'KPIs', CHUNK = 45000;
const HEAD = ['Time','Type','Run ID','Session','Version','Step / action','Provider','Model','Status','ms','Attempt','Fit call','Feedback','Detail','Request','Response','More (long text continues)'];

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
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
function doGet() { return ContentService.createTextOutput('Fit Check log endpoint is running.'); }

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
