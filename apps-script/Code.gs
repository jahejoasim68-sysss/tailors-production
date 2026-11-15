// FYO Production Tracker — backend for the GitHub page.
// Paste this into a Google Sheet: Extensions > Apps Script, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Copy the Web app URL and send it to Claude.

const KEY = '-jBFYlbrWT2DrFKb8jldXl7_'; // must match the page

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}
const STAGES = () => sheet_('Stages', ['Order', 'Stage', 'Updated by', 'Updated at']);
const NOTES  = () => sheet_('Order notes', ['Order', 'Note', 'By', 'At']);
const TEAM   = () => sheet_('Team notes', ['Id', 'Note', 'By', 'At']);

function rows_(sh) { const v = sh.getDataRange().getValues(); v.shift(); return v; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ts_(d) { return d instanceof Date ? d.getTime() : Number(d) || 0; }

function doGet(e) {
  if ((e.parameter.key || '') !== KEY) return json_({ error: 'bad key' });
  const stages = {}, notes = {};
  rows_(STAGES()).forEach(r => { if (r[0]) stages[String(r[0])] = { stage: r[1], by: r[2], at: ts_(r[3]) }; });
  rows_(NOTES()).forEach(r => { if (r[0]) notes[String(r[0])] = { text: r[1], by: r[2], at: ts_(r[3]) }; });
  const team = rows_(TEAM()).filter(r => r[0]).map(r => ({ id: String(r[0]), text: r[1], by: r[2], at: ts_(r[3]) }))
    .sort((a, b) => b.at - a.at).slice(0, 100);
  return json_({ stages, notes, team });
}

function upsert_(sh, id, values) {
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === id) { sh.getRange(i + 1, 1, 1, values.length).setValues([values]); return; }
  sh.appendRow(values);
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    if (b.key !== KEY) return json_({ error: 'bad key' });
    const by = String(b.by || '').slice(0, 60), now = new Date();
    const n = String(b.n || '').replace(/[^0-9]/g, '');
    if (b.action === 'stage' && n && ['new','emb','sti','rdy','dis'].includes(b.stage)) upsert_(STAGES(), n, [n, b.stage, by, now]);
    else if (b.action === 'note' && n) upsert_(NOTES(), n, [n, String(b.text || '').slice(0, 2000), by, now]);
    else if (b.action === 'team' && b.text) TEAM().appendRow([Utilities.getUuid(), String(b.text).slice(0, 2000), by, now]);
    else if (b.action === 'teamdel' && b.id) {
      const sh = TEAM(), v = sh.getDataRange().getValues();
      for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(b.id)) sh.deleteRow(i + 1);
    } else return json_({ error: 'bad request' });
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}
