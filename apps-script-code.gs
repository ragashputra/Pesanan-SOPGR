/**
 * Backend Pesan Makanan & Minuman — Google Apps Script
 * Mendukung: simpan pesanan (POST), cek pesanan (?action=check&id=...), daftar pesanan (?action=list)
 *
 * Kolom sheet "Hasil Pemesanan":
 * A Waktu | B Nama | C Makanan | D Minuman | E Catatan | F Harga Makanan | G Harga Minuman | H Total | I ID
 * (Kolom I "ID" dibuat otomatis bila belum ada. Data lama tanpa ID tetap tampil di daftar.)
 */

var SPREADSHEET_ID = "";            // Kosongkan bila script dibuat dari Extensions > Apps Script di dalam spreadsheet-nya.
var SHEET_NAME     = "Hasil Pemesanan";
var TZ             = "Asia/Jakarta";
var MAX_TOTAL      = 60000;         // batas per orang (rupiah)
var MAX_LIST       = 1000;          // maksimum baris yang dikirim ke daftar
var VERSION        = 3;

var HEADERS = ["Waktu", "Nama", "Makanan", "Minuman", "Catatan", "Harga Makanan", "Harga Minuman", "Total", "ID"];

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
    sh.setFrozenRows(1);
  } else if (String(sh.getRange(1, 9).getValue()).trim() === "") {
    sh.getRange(1, 9).setValue("ID").setFontWeight("bold");
  }
  return sh;
}

function clean_(v, max) {
  v = String(v == null ? "" : v).replace(/[\u0000-\u001f\u007f]+/g, " ").trim();
  if (v.length > max) v = v.substring(0, max);
  return v;
}

function toInt_(v) {
  var n = parseInt(String(v == null ? "" : v).replace(/[^\d-]/g, ""), 10);
  return isNaN(n) ? 0 : n;
}

// Waktu dari sheet -> milidetik. Bisa berupa Date asli atau teks "dd/MM/yyyy HH:mm:ss".
function toMs_(v) {
  if (v instanceof Date) return v.getTime();
  if (typeof v === "number") return v;
  var s = String(v || "").trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ ,]+(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?)?$/);
  if (m) {
    // anggap waktu Jakarta (UTC+7)
    return Date.UTC(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)) - 7 * 3600 * 1000;
  }
  var t = Date.parse(s);
  return isNaN(t) ? 0 : t;
}

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    var action = String(p.action || "").toLowerCase();

    if (action === "list") {
      var sh = sheet_();
      var last = sh.getLastRow();
      var orders = [];
      if (last >= 2) {
        var start = Math.max(2, last - MAX_LIST + 1);
        var rows = sh.getRange(start, 1, last - start + 1, 8).getValues();
        for (var i = rows.length - 1; i >= 0; i--) {        // terbaru di atas
          var r = rows[i];
          if (!String(r[1]).trim() && !String(r[2]).trim() && !String(r[3]).trim()) continue; // lewati baris kosong
          var total = toInt_(r[7]);
          if (!total) total = toInt_(r[5]) + toInt_(r[6]);
          orders.push({
            t: toMs_(r[0]),
            nama: String(r[1]),
            makanan: String(r[2]),
            minuman: String(r[3]),
            catatan: String(r[4]),
            total: total
          });
        }
      }
      return json_({ ok: true, v: VERSION, orders: orders });
    }

    if (action === "check") {
      var id = clean_(p.id, 80);
      if (!id) return json_({ ok: false, error: "no id" });
      return json_({ ok: true, found: exists_(sheet_(), id) });
    }

    return json_({ ok: true, v: VERSION, service: "pesan-makanan" });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function exists_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return false;
  var ids = sh.getRange(2, 9, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === id) return true;
  return false;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: "busy" });
  }
  try {
    var p = (e && e.parameter) || {};
    var id      = clean_(p.id, 80);
    var nama    = clean_(p.nama, 60);
    var makanan = clean_(p.makanan, 80);
    var minuman = clean_(p.minuman, 80);
    var catatan = clean_(p.catatan, 300);
    var hm      = toInt_(p.hargaMakanan);
    var hd      = toInt_(p.hargaMinuman);
    var total   = hm + hd;                                  // dihitung ulang di server

    if (!id || nama.length < 2 || !makanan || !minuman) return json_({ ok: false, error: "data tidak lengkap" });
    if (hm < 0 || hd < 0 || total <= 0 || total > MAX_TOTAL) return json_({ ok: false, error: "total melebihi batas" });

    var sh = sheet_();
    if (exists_(sh, id)) return json_({ ok: true, duplicate: true });   // kirim ulang tidak menggandakan baris

    var row = sh.getLastRow() + 1;
    // kolom teks dipaksa sebagai teks agar teks berawalan = + - @ tidak dibaca sebagai rumus
    sh.getRange(row, 2, 1, 4).setNumberFormat("@");
    sh.getRange(row, 9).setNumberFormat("@");
    sh.getRange(row, 1, 1, 9).setValues([[new Date(), nama, makanan, minuman, catatan, hm, hd, total, id]]);
    sh.getRange(row, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sh.getRange(row, 6, 1, 3).setNumberFormat('"Rp"#,##0');
    SpreadsheetApp.flush();

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/** Jalankan sekali dari editor (tombol Run) untuk memberi izin akses spreadsheet. */
function setup() {
  var sh = sheet_();
  Logger.log("OK, sheet: " + sh.getName() + ", baris: " + sh.getLastRow());
}
