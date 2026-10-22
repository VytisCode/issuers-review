/**
 * ISIN sąrašo saugykla dashboard'ui (Google Apps Script, nemokama).
 * Įdiegimas - README skiltis "ISIN sąrašas".
 *  1. Google lentelėje: Extensions → Apps Script, įklijuokite šį kodą, Save.
 *  2. (Nebūtina) Project Settings → Script properties: PIN = bendras kodas, kurį žinos kolegos.
 *  3. Deploy → New deployment → Web app; Execute as: Me; Who has access: Anyone → Deploy, nukopijuokite URL.
 *  4. URL įrašykite į config/settings.yaml "isin_store_url".
 * Lapai: "ISIN" (isin, note, added, added_by) ir "Log" (laikas, kas, pridėta, išimta).
 */
function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(header);
  return sh;
}

function read_() {
  var sh = sheet_("ISIN", ["isin", "note", "added", "added_by"]);
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
  return rows.filter(function (r) { return r[0]; }).map(function (r) {
    return { isin: String(r[0]).trim().toUpperCase(), note: String(r[1] || ""), added: r[2] ? Utilities.formatDate(new Date(r[2]), "UTC", "yyyy-MM-dd") : "", added_by: String(r[3] || "") };
  });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var log = sheet_("Log", ["time", "by", "added", "removed"]);
  var last = log.getLastRow() > 1 ? log.getRange(log.getLastRow(), 1, 1, 4).getValues()[0] : null;
  return out_({ ok: true, isins: read_(), pin_required: !!PropertiesService.getScriptProperties().getProperty("PIN"),
                last_change: last ? { time: new Date(last[0]).toISOString(), by: last[1], added: last[2], removed: last[3] } : null });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var req = JSON.parse(e.postData.contents || "{}");
    var pin = PropertiesService.getScriptProperties().getProperty("PIN");
    if (pin && String(req.pin || "") !== pin) return out_({ ok: false, error: "pin" });
    var re = /^[A-Z]{2}[A-Z0-9]{9}[0-9]$/;
    var notes = req.notes || {};
    var want = [];
    (req.isins || []).forEach(function (x) {
      x = String(x).trim().toUpperCase();
      if (re.test(x) && want.indexOf(x) < 0) want.push(x);
    });
    if (want.length > 500) return out_({ ok: false, error: "too_many" });
    var cur = read_(), curIds = cur.map(function (r) { return r.isin; });
    var by = String(req.by || "").slice(0, 60);
    var now = new Date();
    var rows = want.map(function (x) {
      var old = cur[curIds.indexOf(x)];
      return [x, notes[x] !== undefined ? String(notes[x]).slice(0, 200) : (old ? old.note : ""), old && old.added ? old.added : now, old ? old.added_by : by];
    });
    var sh = sheet_("ISIN", ["isin", "note", "added", "added_by"]);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
    var added = want.filter(function (x) { return curIds.indexOf(x) < 0; });
    var removed = curIds.filter(function (x) { return want.indexOf(x) < 0; });
    if (added.length || removed.length) sheet_("Log", ["time", "by", "added", "removed"]).appendRow([now, by, added.join(" "), removed.join(" ")]);
    return out_({ ok: true, added: added, removed: removed, count: want.length });
  } finally {
    lock.releaseLock();
  }
}
