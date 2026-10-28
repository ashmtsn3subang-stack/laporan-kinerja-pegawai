/**
 * Laporan Kinerja PPPK - server Google Apps Script
 *
 * Fungsi: login username dan password, pendaftaran guru dengan kode pendaftaran,
 * dan penyimpanan data tiap guru di Google Sheets.
 *
 * Cara pasang singkat (rincian lengkap ada di panduan):
 * 1. Buat Google Sheets baru, buka Extensions > Apps Script.
 * 2. Tempel file ini sebagai Code.gs, buat file HTML bernama Index (isi dari Index.html),
 *    dan isi appsscript.json (Project Settings > Show appsscript.json manifest file).
 * 3. Jalankan fungsi setup() satu kali dan setujui izin yang diminta.
 * 4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 * 5. Bagikan tautan web app dan kode pendaftaran kepada guru.
 */

var SHEET_USERS = 'Pengguna';
var SHEET_DATA = 'Data';
var USER_HEADERS = ['username', 'salt', 'hash', 'nama', 'peran', 'aktif', 'dibuat', 'login_terakhir'];
var DATA_HEADERS = ['username', 'kunci', 'bagian', 'isi', 'diperbarui'];
var CHUNK = 45000;            // batas sel Google Sheets 50.000 karakter
var SESSION_DAYS = 14;
var HASH_ROUNDS = 400;
var MAX_FAILS = 5;
var LOCK_SECONDS = 600;
var MEDIA_KEYS = ['logo', 'ttdPegawai', 'ttdKepala', 'stempel'];
var MAX_MAIN = 400000;        // batas ukuran data utama satu guru (karakter)
var MAX_IMG = 220000;         // batas ukuran satu gambar (karakter data URL)

/* ================= Pemasangan ================= */

/** Jalankan sekali dari editor. Membuat lembar, rahasia server, dan kode pendaftaran. */
function setup() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEET_USERS, USER_HEADERS);
  ensureSheet_(ss, SHEET_DATA, DATA_HEADERS);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', randomString_(48));
  if (!props.getProperty('PEPPER')) props.setProperty('PEPPER', randomString_(48));
  if (!props.getProperty('KODE_DAFTAR')) props.setProperty('KODE_DAFTAR', randomString_(8).toUpperCase());
  var msg = 'Pemasangan selesai.\nKode pendaftaran guru: ' + props.getProperty('KODE_DAFTAR') +
    '\nGuru pertama yang mendaftar otomatis menjadi admin.\nUbah kode lewat menu Laporan Kinerja > Ubah kode pendaftaran.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* dijalankan di luar spreadsheet */ }
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Laporan Kinerja')
      .addItem('Tambah guru', 'menuTambahGuru')
      .addItem('Atur ulang sandi guru', 'menuResetSandi')
      .addItem('Nonaktifkan atau aktifkan guru', 'menuToggleAktif')
      .addItem('Lihat kode pendaftaran', 'menuLihatKode')
      .addItem('Ubah kode pendaftaran', 'menuUbahKode')
      .addToUi();
  } catch (e) { /* abaikan */ }
}

function menuTambahGuru() {
  var ui = SpreadsheetApp.getUi();
  var u = ui.prompt('Tambah guru', 'Username (huruf kecil, angka, titik, minus, garis bawah; 3-30 karakter):', ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK) return;
  var n = ui.prompt('Tambah guru', 'Nama lengkap guru:', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  var p = ui.prompt('Tambah guru', 'Sandi awal (minimal 8 karakter):', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  var r = createUser_(u.getResponseText(), p.getResponseText(), n.getResponseText());
  ui.alert(r.ok ? 'Guru ditambahkan.' : r.error);
}

function menuResetSandi() {
  var ui = SpreadsheetApp.getUi();
  var u = ui.prompt('Atur ulang sandi', 'Username guru:', ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK) return;
  var p = ui.prompt('Atur ulang sandi', 'Sandi baru (minimal 8 karakter):', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  var r = setPassword_(u.getResponseText(), p.getResponseText());
  ui.alert(r.ok ? 'Sandi diganti. Sampaikan sandi baru kepada guru.' : r.error);
}

function menuToggleAktif() {
  var ui = SpreadsheetApp.getUi();
  var u = ui.prompt('Aktif atau nonaktif', 'Username guru:', ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK) return;
  var f = findUser_(normUser_(u.getResponseText()));
  if (!f) { ui.alert('Username tidak ditemukan.'); return; }
  var now = f.row[5] === false || String(f.row[5]).toLowerCase() === 'false';
  f.sheet.getRange(f.index, 6).setValue(now);
  ui.alert(now ? 'Akun diaktifkan.' : 'Akun dinonaktifkan. Guru tidak bisa masuk lagi.');
}

function menuLihatKode() {
  SpreadsheetApp.getUi().alert('Kode pendaftaran: ' + PropertiesService.getScriptProperties().getProperty('KODE_DAFTAR'));
}

function menuUbahKode() {
  var ui = SpreadsheetApp.getUi();
  var k = ui.prompt('Ubah kode pendaftaran', 'Kode baru (minimal 6 karakter):', ui.ButtonSet.OK_CANCEL);
  if (k.getSelectedButton() !== ui.Button.OK) return;
  var t = k.getResponseText().trim();
  if (t.length < 6) { ui.alert('Kode terlalu pendek.'); return; }
  PropertiesService.getScriptProperties().setProperty('KODE_DAFTAR', t);
  ui.alert('Kode pendaftaran diganti.');
}

/* ================= Halaman web ================= */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Laporan Kinerja PPPK')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOW_ALL);
}

/* ================= API yang dipanggil halaman ================= */

function apiRegister(username, password, nama, kode) {
  try {
    var expected = PropertiesService.getScriptProperties().getProperty('KODE_DAFTAR') || '';
    if (!expected) return fail_('Server belum dipasang. Jalankan setup() terlebih dahulu.');
    var key = 'reg:' + 'all';
    if (isLocked_(key)) return fail_('Terlalu banyak percobaan. Coba lagi beberapa menit lagi.');
    if (!safeEqual_(String(kode || '').trim().toUpperCase(), expected.toUpperCase())) {
      addFail_(key);
      return fail_('Kode pendaftaran salah.');
    }
    var r = createUser_(username, password, nama);
    if (!r.ok) return r;
    return session_(r.user);
  } catch (e) {
    return fail_('Terjadi kesalahan di server. Coba lagi.');
  }
}

function apiLogin(username, password) {
  try {
    var u = normUser_(username);
    if (!u || !password) return fail_('Isi username dan sandi.');
    var key = 'login:' + u;
    if (isLocked_(key)) return fail_('Akun dikunci sementara karena terlalu banyak percobaan. Coba lagi dalam 10 menit.');
    var f = findUser_(u);
    var ok = false;
    if (f) {
      var salt = String(f.row[1]), hash = String(f.row[2]);
      ok = safeEqual_(hashPassword_(password, salt), hash);
    } else {
      hashPassword_(password, 'x'); // samakan waktu respons
    }
    if (!ok) { addFail_(key); return fail_('Username atau sandi salah.'); }
    if (f.row[5] === false || String(f.row[5]).toLowerCase() === 'false') return fail_('Akun ini dinonaktifkan. Hubungi admin.');
    CacheService.getScriptCache().remove(key);
    f.sheet.getRange(f.index, 8).setValue(new Date());
    return session_(u);
  } catch (e) {
    return fail_('Terjadi kesalahan di server. Coba lagi.');
  }
}

function apiLoad(token) {
  try {
    var a = auth_(token);
    if (!a.ok) return a;
    var out = { ok: true, user: a.user, nama: a.nama, peran: a.peran, data: readValue_(a.user, 'main'), media: {} };
    MEDIA_KEYS.forEach(function (k) {
      var v = readValue_(a.user, 'img-' + k);
      if (v) { try { out.media[k] = JSON.parse(v); } catch (e) { out.media[k] = null; } }
    });
    return out;
  } catch (e) {
    return fail_('Data belum bisa dibaca. Coba lagi.');
  }
}

/** main: teks JSON data utama. media: objek {kunci: {url,w,h} atau null}. */
function apiSave(token, main, media) {
  try {
    var a = auth_(token);
    if (!a.ok) return a;
    if (typeof main !== 'string' || main.length > MAX_MAIN) return fail_('Data terlalu besar untuk disimpan.');
    try { JSON.parse(main); } catch (e) { return fail_('Data tidak valid.'); }
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      writeValue_(a.user, 'main', main);
      media = media || {};
      MEDIA_KEYS.forEach(function (k) {
        if (!Object.prototype.hasOwnProperty.call(media, k)) return;
        var m = media[k];
        if (m && m.url) {
          var s = JSON.stringify({ url: String(m.url), w: Number(m.w) || 0, h: Number(m.h) || 0 });
          if (s.length > MAX_IMG + 100 || String(m.url).indexOf('data:image/') !== 0) throw new Error('gambar');
          writeValue_(a.user, 'img-' + k, s);
        } else {
          writeValue_(a.user, 'img-' + k, null);
        }
      });
    } finally {
      lock.releaseLock();
    }
    return { ok: true };
  } catch (e) {
    return fail_(String(e && e.message) === 'gambar' ? 'Gambar tidak valid atau terlalu besar.' : 'Data belum tersimpan. Coba lagi.');
  }
}

function apiChangePassword(token, oldPass, newPass) {
  try {
    var a = auth_(token);
    if (!a.ok) return a;
    var f = findUser_(a.user);
    if (!safeEqual_(hashPassword_(String(oldPass || ''), String(f.row[1])), String(f.row[2]))) return fail_('Sandi lama salah.');
    return setPassword_(a.user, newPass);
  } catch (e) {
    return fail_('Sandi belum bisa diganti. Coba lagi.');
  }
}

/* ================= Pengguna dan sesi ================= */

function normUser_(u) { return String(u || '').trim().toLowerCase(); }

function createUser_(username, password, nama) {
  var u = normUser_(username);
  if (!/^[a-z0-9._-]{3,30}$/.test(u)) return fail_('Username 3-30 karakter: huruf kecil, angka, titik, minus, atau garis bawah.');
  if (String(password || '').length < 8) return fail_('Sandi minimal 8 karakter.');
  var name = String(nama || '').trim();
  if (!name) return fail_('Isi nama lengkap.');
  if (name.length > 100) return fail_('Nama terlalu panjang.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (findUser_(u)) return fail_('Username sudah dipakai. Pilih username lain.');
    var sheet = userSheet_();
    var first = sheet.getLastRow() < 2;
    var salt = randomString_(24);
    sheet.appendRow([u, salt, hashPassword_(password, salt), name, first ? 'admin' : 'guru', true, new Date(), '']);
    return { ok: true, user: u };
  } finally {
    lock.releaseLock();
  }
}

function setPassword_(username, newPass) {
  if (String(newPass || '').length < 8) return fail_('Sandi minimal 8 karakter.');
  var f = findUser_(normUser_(username));
  if (!f) return fail_('Username tidak ditemukan.');
  var salt = randomString_(24);
  f.sheet.getRange(f.index, 2, 1, 2).setValues([[salt, hashPassword_(newPass, salt)]]);
  CacheService.getScriptCache().remove('login:' + f.row[0]);
  return { ok: true };
}

function findUser_(u) {
  if (!u) return null;
  var sheet = userSheet_(), last = sheet.getLastRow();
  if (last < 2) return null;
  var rows = sheet.getRange(2, 1, last - 1, USER_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === u) return { sheet: sheet, index: i + 2, row: rows[i] };
  }
  return null;
}

function session_(user) {
  var f = findUser_(user);
  return { ok: true, token: makeToken_(user), user: user, nama: String(f.row[3]), peran: String(f.row[4]) };
}

function auth_(token) {
  var u = verifyToken_(token);
  if (!u) return fail_('Sesi berakhir. Masuk lagi.', 'sesi');
  var f = findUser_(u);
  if (!f || f.row[5] === false || String(f.row[5]).toLowerCase() === 'false') return fail_('Akun tidak aktif.', 'sesi');
  return { ok: true, user: u, nama: String(f.row[3]), peran: String(f.row[4]) };
}

function makeToken_(user) {
  var exp = Date.now() + SESSION_DAYS * 86400000;
  var payload = Utilities.base64EncodeWebSafe(user + '|' + exp);
  return payload + '.' + sign_(payload);
}

function verifyToken_(token) {
  if (typeof token !== 'string' || token.indexOf('.') < 1) return null;
  var parts = token.split('.');
  if (parts.length !== 2 || !safeEqual_(sign_(parts[0]), parts[1])) return null;
  var raw;
  try { raw = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); } catch (e) { return null; }
  var i = raw.lastIndexOf('|');
  if (i < 1) return null;
  if (Number(raw.slice(i + 1)) < Date.now()) return null;
  return raw.slice(0, i);
}

function sign_(text) {
  var secret = PropertiesService.getScriptProperties().getProperty('SECRET') || '';
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text, secret));
}

function hashPassword_(password, salt) {
  var pepper = PropertiesService.getScriptProperties().getProperty('PEPPER') || '';
  var h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pepper + ':' + salt + ':' + password, Utilities.Charset.UTF_8);
  var saltBytes = Utilities.newBlob(salt).getBytes();
  for (var i = 0; i < HASH_ROUNDS; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h.concat(saltBytes));
  }
  return Utilities.base64EncodeWebSafe(h);
}

function isLocked_(key) { return Number(CacheService.getScriptCache().get('fail:' + key) || 0) >= MAX_FAILS; }
function addFail_(key) {
  var c = CacheService.getScriptCache(), k = 'fail:' + key;
  c.put(k, String(Number(c.get(k) || 0) + 1), LOCK_SECONDS);
}

function safeEqual_(a, b) {
  a = String(a); b = String(b);
  var diff = a.length ^ b.length;
  for (var i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

function fail_(msg, code) { return { ok: false, error: msg, code: code || '' }; }

function randomString_(n) {
  var chars = 'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', out = '';
  var u = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, u + Date.now() + Math.random());
  while (out.length < n) {
    for (var i = 0; i < bytes.length && out.length < n; i++) out += chars.charAt((bytes[i] + 256) % chars.length);
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  }
  return out;
}

/* ================= Penyimpanan di Sheets ================= */

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function userSheet_() { return ensureSheet_(SpreadsheetApp.getActive(), SHEET_USERS, USER_HEADERS); }
function dataSheet_() { return ensureSheet_(SpreadsheetApp.getActive(), SHEET_DATA, DATA_HEADERS); }

function readValue_(user, key) {
  var sh = dataSheet_(), last = sh.getLastRow();
  if (last < 2) return null;
  var rows = sh.getRange(2, 1, last - 1, 4).getValues(), parts = [];
  rows.forEach(function (r) { if (String(r[0]) === user && String(r[1]) === key) parts.push({ n: Number(r[2]), t: String(r[3]) }); });
  if (!parts.length) return null;
  parts.sort(function (a, b) { return a.n - b.n; });
  return parts.map(function (p) { return p.t; }).join('');
}

/** value null menghapus. Dipanggil di dalam kunci skrip. */
function writeValue_(user, key, value) {
  var sh = dataSheet_(), last = sh.getLastRow();
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][0]) === user && String(rows[i][1]) === key) sh.deleteRow(i + 2);
    }
  }
  if (value == null || value === '') return;
  var now = new Date(), out = [];
  for (var p = 0, n = 0; p < value.length; p += CHUNK, n++) out.push([user, key, n, value.substr(p, CHUNK), now]);
  sh.getRange(sh.getLastRow() + 1, 1, out.length, 5).setNumberFormat('@').setValues(out);
}
