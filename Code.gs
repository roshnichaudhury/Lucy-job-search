// Lucy interaction logger. Paste into the Lucy Interaction Log sheet: Extensions > Apps Script.
// Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone. Copy the web app URL.
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = JSON.parse(e.postData.contents || "{}");
  var events = (data.events || []).slice(0, 200);
  var now = new Date();
  var rows = events.map(function (ev) {
    return [now, s(ev.session), s(ev.t), s(ev.kind), s(ev.step), s(ev.value), s(ev.tone)];
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  return ContentService.createTextOutput("ok");
}
function s(v) { return String(v == null ? "" : v).slice(0, 1000); }
