// V91.StokOrd - jembatan Google Sheet
// Tempel di Apps Script, lalu Terapkan > Penerapan baru > Aplikasi web
// Jalankan sebagai: Saya | Siapa yang punya akses: Siapa saja
var KEY = 'GANTI-KUNCI';
var SS_ID = '12XDpFeyekAlEUQaqGhrR_RvpgU5xFafi3_fyRofco1s';

function doPost(e) {
  var out;
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.key !== KEY) throw new Error('Kunci rahasia tidak cocok');
    if (d.action === 'ping') {
      out = { ok: true };
    } else {
      var lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        var ss = SpreadsheetApp.openById(SS_ID);
        if (d.action === 'push') {
          Object.keys(d.tabs).forEach(function (name) {
            var sh = ss.getSheetByName(name) || ss.insertSheet(name);
            var v = d.tabs[name];
            sh.clearContents();
            if (v.length && v[0].length) sh.getRange(1, 1, v.length, v[0].length).setValues(v);
          });
          if (d.info && d.info.length) {
            var s2 = ss.getSheetByName('Info') || ss.insertSheet('Info');
            s2.clearContents();
            s2.getRange(1, 1, d.info.length, d.info[0].length).setValues(d.info);
          }
          out = { ok: true };
        } else if (d.action === 'pull') {
          var tabs = {};
          ['Bibit', 'Kemasan', 'Mitra', 'Transaksi'].forEach(function (name) {
            var sh = ss.getSheetByName(name);
            tabs[name] = (sh && sh.getLastRow() > 0) ? sh.getDataRange().getValues() : [];
          });
          out = { ok: true, tabs: tabs };
        } else {
          throw new Error('Aksi tidak dikenal');
        }
      } finally {
        lock.releaseLock();
      }
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('V91.StokOrd jembatan aktif');
}
