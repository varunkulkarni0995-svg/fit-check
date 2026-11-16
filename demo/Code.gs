/**
 * Fit Check live demo: a small server (Google Apps Script) for the demo app.
 * - Holds the team's FREE-TIER keys in Script properties (GROQ_KEY, GEMINI_KEY). They never reach the browser.
 * - Forwards AI calls only to Groq and Gemini.
 * - Keeps the shared count of live checks (limit 100) and a daily cap on AI calls.
 * Setup: script.google.com > New project > paste this > Project Settings > Script properties (GROQ_KEY, GEMINI_KEY)
 *        > Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone > Deploy > copy the /exec URL.
 */
const DEMO_LIMIT = 100, DAILY_CALL_CAP = 1500;
const HOSTS = {'api.groq.com': 'GROQ_KEY', 'generativelanguage.googleapis.com': 'GEMINI_KEY'};
const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet(e) {
  const p = (e && e.parameter) || {}, props = PropertiesService.getScriptProperties();
  if (p.action === 'demo_status' || p.action === 'demo_take') {
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      let used = +(props.getProperty('demo_used') || 0); const ok = used < DEMO_LIMIT;
      if (p.action === 'demo_take' && ok) { used++; props.setProperty('demo_used', String(used)); }
      return json_({used: used, limit: DEMO_LIMIT, ok: ok});
    } finally { lock.releaseLock(); }
  }
  return ContentService.createTextOutput('Fit Check demo server is running.');
}

function doPost(e) {
  let d; try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({status: 400, text: '{"error":{"message":"Bad request"}}'}); }
  if (d.type !== 'llm') return json_({status: 400, text: '{"error":{"message":"Unknown request"}}'});
  const m = String(d.url || '').match(/^https:\/\/([^\/]+)\//), host = m && m[1];
  if (!host || !HOSTS[host]) return json_({status: 400, text: '{"error":{"message":"Host not allowed"}}'});
  const props = PropertiesService.getScriptProperties(), key = props.getProperty(HOSTS[host]);
  if (!key) return json_({status: 401, text: '{"error":{"message":"Team key not set up"}}'});
  const day = 'calls_' + Utilities.formatDate(new Date(), 'America/New_York', 'yyyyMMdd'), n = +(props.getProperty(day) || 0);
  if (n >= DAILY_CALL_CAP) return json_({status: 429, text: '{"error":{"message":"Team daily limit reached (per day)"}}'});
  props.setProperty(day, String(n + 1));
  const headers = host === 'api.groq.com' ? {Authorization: 'Bearer ' + key} : {'x-goog-api-key': key};
  const r = UrlFetchApp.fetch(d.url, {method: 'post', contentType: 'application/json', headers: headers, payload: JSON.stringify(d.body || {}), muteHttpExceptions: true});
  return json_({status: r.getResponseCode(), text: r.getContentText()});
}
