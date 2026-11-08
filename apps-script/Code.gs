/**
 * Akato Sangha - shared accounts backend (Google Apps Script).
 * 1. Create a Google Sheet. Extensions > Apps Script. Paste this file.
 * 2. Change KEY below to a private committee key (share it only with committee members).
 * 3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Copy the Web app URL.
 */
var KEY = 'aks-JNszl1kLusD-hsdw';
var SHEET_ID = '1xSuUvZiwacxyrGZY5Y0ftYtu3fhPARup_7qiXdF8am8';
var SHEET = 'data';
var CHUNK = 45000, NCH = 20;

function sh_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(['col','id','u','su','deleted'].concat(Array.apply(null, Array(NCH)).map(function(_, i){ return 'json' + i; }))); s.setFrozenRows(1); }
  return s;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return out_({ ok: true, app: 'akato-sangha' }); }

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ error: 'bad request' }); }
  if (KEY === 'CHANGE-ME') return out_({ error: 'key not set in script' });
  if (req.key !== KEY) return out_({ error: 'bad key' });
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var s = sh_();
    var last = s.getLastRow();
    var data = last > 1 ? s.getRange(2, 1, last - 1, 5 + NCH).getValues() : [];
    var idx = {};
    data.forEach(function (r, i) { idx[r[0] + '|' + r[1]] = i; });
    var now = Date.now();
    var since = +req.since || 0;
    var ops = req.ops || [];
    var touched = {};
    ops.forEach(function (op) {
      if (!op || !op.c || !op.id) return;
      var k = op.c + '|' + op.id, u = +op.u || now;
      var json = op.d === null ? '' : JSON.stringify(op.d);
      var row = [op.c, op.id, u, now, op.d === null ? 1 : 0];
      for (var i = 0; i < NCH; i++) row.push(json.substr(i * CHUNK, CHUNK));
      if (k in idx) {
        var cur = data[idx[k]];
        if (+cur[2] > u) return; // newer edit already stored
        data[idx[k]] = row; touched[idx[k]] = true;
      } else {
        data.push(row); idx[k] = data.length - 1; touched[idx[k]] = true;
      }
    });
    var ks = Object.keys(touched).map(Number);
    ks.forEach(function (i) {
      s.getRange(i + 2, 1, 1, 5 + NCH).setValues([data[i]]);
    });
    var rows = [];
    data.forEach(function (r) {
      if (+r[3] > since) {
        var j = '';
        for (var i = 0; i < NCH; i++) j += r[5 + i];
        rows.push({ c: r[0], id: String(r[1]), u: +r[2], d: (+r[4] === 1 || !j) ? null : JSON.parse(j) });
      }
    });
    return out_({ now: now, rows: rows });
  } catch (err) {
    return out_({ error: String(err) });
  } finally { lock.releaseLock(); }
}
