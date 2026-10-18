/**
 * MONITORING OUTSTANDING CROSSDOCKING — Google Apps Script
 * Database: Google Sheets | Auth: salted+peppered SHA-256, token sesi di CacheService
 */

var TZ = 'Asia/Jakarta';
var CFG = {
  TTL: 21600,            // sesi 6 jam (detik)
  MAX_FAIL: 5,           // maks gagal login
  LOCK_SEC: 900,         // kunci 15 menit
  KATEGORI: ['Inbound', 'Outbound'],
  UKURAN: ['20 ft', '40 ft', '40 HC', '45 ft', 'Wingbox', 'CDD / Fuso', 'Blindvan'],
  JENIS_KARGO: ['Sparepart'],
  STATUS: ['Belum Dikerjakan', 'On Progress', 'Selesai'],
  STATUS_PENGAWAS: ['On Progress', 'Selesai'],
  HASIL: ['Complete', 'Overload', 'Damage', 'Lainnya'],
  AREA: ['Phase 2', 'Phase 4', 'Phase 8'],
  FOLLOWUP_MS: 5 * 60 * 1000,   // jeda 5 menit sebelum CS mengirim pesan penutup
  ROLES: ['ADMIN', 'LEADER', 'PENGAWAS'],
  DEFAULT_ADMIN: { username: 'admin' }   // password awal dibuat acak oleh adminPw_() dan disimpan di Script Properties
};

var SCHEMA = {
  Users: ['ID', 'Username', 'Hash', 'Salt', 'Nama', 'Role', 'Tipe', 'CustomerIDs', 'Aktif', 'Dibuat', 'LoginTerakhir'],
  Customers: ['ID', 'Nama', 'TampilPublik', 'Aktif'],
  Checkers: ['ID', 'Nama', 'Aktif'],
  Jobs: ['ID', 'Tanggal', 'CustomerID', 'Customer', 'Kategori', 'Status', 'Detail', 'Ukuran', 'NoContainer', 'Closing',
    'Checker', 'PengawasUser', 'PengawasNama', 'LeaderUser', 'LeaderNama', 'Hasil', 'Catatan', 'Dibuat', 'Diupdate', 'DiupdateOleh', 'Selesai', 'JmlArmada', 'JenisKargo', 'Area'],
  Containers: ['JobID', 'NoContainer', 'Done', 'DoneAt', 'DoneBy'],
  History: ['Waktu', 'JobID', 'NoContainer', 'Customer', 'Dari', 'Ke', 'Hasil', 'Catatan', 'Oleh', 'Peran'],
  Chats: ['ChatID', 'Waktu', 'Nama', 'Customer', 'NoContainer', 'Status', 'LeaderUser', 'PengawasUser', 'Diupdate', 'DiupdateOleh', 'FollowUpAt'],
  ChatMsgs: ['ChatID', 'Waktu', 'Dari', 'Pengirim', 'Pesan'],
  AuditLog: ['Waktu', 'User', 'Role', 'Aksi', 'Detail']
};

/* ============================ ENTRY ============================ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Job Outstanding')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================ SETUP ============================ */
// Password awal: dibuat acak sekali, disimpan di Script Properties (BUKAN di kode), lalu dilihat lewat View > Logs.
function adminPw_() {
  var p = PropertiesService.getScriptProperties(), k = 'ADMIN_INIT_PW', v = p.getProperty(k);
  if (!v) { v = 'Adm' + Utilities.getUuid().replace(/-/g, '').slice(0, 10) + '9'; p.setProperty(k, v); }
  return v;
}

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SSID'));
  var props = PropertiesService.getScriptProperties();
  props.setProperty('SSID', ss.getId());
  if (!props.getProperty('PEPPER')) props.setProperty('PEPPER', Utilities.getUuid() + Utilities.getUuid());
  Object.keys(SCHEMA).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var h = SCHEMA[name];
    sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#E7E2FB');
    sh.setFrozenRows(1);
  });
  Object.keys(SCHEMA).forEach(function (n) { bump_(n); });
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  if (readAll_('Users').length === 0) {
    createUser_({ username: CFG.DEFAULT_ADMIN.username, nama: 'Administrator', role: 'ADMIN', password: adminPw_(), aktif: 'Y' });
    audit_({ Username: 'system', Role: 'SYSTEM' }, 'SETUP', 'Admin default dibuat');
  }
  Logger.log('Setup selesai. Login awal: ' + CFG.DEFAULT_ADMIN.username + ' / ' + adminPw_() + '  (segera ganti setelah login)');
}

function seedContoh() {
  setupDatabase();
  var cs = ['PT Alfa Retail', 'PT Beta Foods', 'PT Gamma Fashion', 'PT Delta Electronics'];
  var ids = cs.map(function (n, i) { return saveMaster_('Customers', { nama: n, tampilPublik: i < 3 ? 'Y' : 'N', aktif: 'Y' }, 'C'); });
  ['Checker Demo 1', 'Checker Demo 2', 'Checker Demo 3'].forEach(function (n) { saveMaster_('Checkers', { nama: n, aktif: 'Y' }, 'K'); });
  var mk = function (u, n, r, t, c) { if (!findUser_(u)) createUser_({ username: u, nama: n, role: r, tipe: t || '', customerIDs: c || '', password: adminPw_(), aktif: 'Y' }); };
  mk('leader1', 'Leader Demo 1', 'LEADER', 'Project', ids.slice(0, 2).join(','));
  mk('leader2', 'Leader Demo 2', 'LEADER', 'Non Project', ids.slice(1).join(','));
  mk('pengawas1', 'Pengawas Demo 1', 'PENGAWAS'); mk('pengawas2', 'Pengawas Demo 2', 'PENGAWAS');
  mk('pengawas3', 'Pengawas Demo 3', 'PENGAWAS'); mk('pengawas4', 'Pengawas Demo 4', 'PENGAWAS');
  Logger.log('Password semua user demo (leader1, leader2, pengawas1-4): ' + adminPw_());
}

function resetAdminPassword() {
  var u = findUser_(CFG.DEFAULT_ADMIN.username);
  if (!u) throw new Error('User admin tidak ada');
  u.Salt = Utilities.getUuid(); u.Hash = hash_(adminPw_(), u.Salt); u.Aktif = 'Y'; write_('Users', u);
}

/* ============================ HELPERS DB ============================ */
var SS_MEMO = null;
function ss_() {
  if (SS_MEMO) return SS_MEMO;
  var id = PropertiesService.getScriptProperties().getProperty('SSID');
  SS_MEMO = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  return SS_MEMO;
}
function sheet_(n) {
  var s = ss_().getSheetByName(n);
  var h = SCHEMA[n];
  if (!s && h && ss_().getSheetByName('Jobs')) {
    s = ss_().insertSheet(n);
    s.getRange(1, 1, s.getMaxRows(), h.length).setNumberFormat('@');
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#E7E2FB');
    s.setFrozenRows(1);
  }
  if (!s) throw new Error('Database belum diinisialisasi. Jalankan setupDatabase() di editor.');
  if (h && s.getLastColumn() < h.length) {
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#E7E2FB');
  }
  return s;
}
var SS_TZ = '';
// Zona waktu spreadsheet. Teks tanggal yang ditulis ke sheet dibaca Google Sheets memakai zona ini.
function sheetTz_() {
  if (!SS_TZ) { try { SS_TZ = ss_().getSpreadsheetTimeZone() || TZ; } catch (e) { SS_TZ = TZ; } }
  return SS_TZ;
}
// Jika sel terbaca sebagai Date (Sheets otomatis mengubah teks tanggal), format ulang dengan zona SPREADSHEET
// supaya jam kembali sama persis dengan yang ditulis (waktu WIB), tidak bergeser.
function str_(v) { if (v instanceof Date) return Utilities.formatDate(v, sheetTz_(), 'yyyy-MM-dd HH:mm'); return v == null ? '' : String(v); }
var MEMO = {};
function ver_(n) {
  var c = CacheService.getScriptCache(), v = c.get('v_' + n);
  if (!v) { v = String(Date.now()); c.put('v_' + n, v, 21600); }
  return v;
}
function bump_(n) { delete MEMO[n]; CacheService.getScriptCache().put('v_' + n, String(Date.now()) + Math.random().toString(36).substr(2, 4), 21600); }
function normStatus_(s) { return s === 'Dikerjakan' ? 'On Progress' : (s || 'Belum Dikerjakan'); }
function readSheet_(n) {
  var v = sheet_(n).getDataRange().getValues(), h = SCHEMA[n], out = [];
  for (var i = 1; i < v.length; i++) {
    var o = { _row: i + 1 };
    for (var c = 0; c < h.length; c++) o[h[c]] = str_(v[i][c]);
    if (n === 'Jobs') o.Status = normStatus_(o.Status);
    if (o[h[0]] !== '') out.push(o);
  }
  return out;
}
function readAll_(n) {
  if (MEMO[n]) return MEMO[n];
  var ca = CacheService.getScriptCache(), v = ver_(n), mk = 'm_' + n + v, meta = ca.get(mk);
  if (meta) {
    var keys = [], i; for (i = 0; i < Number(meta); i++) keys.push('d_' + n + v + '_' + i);
    var got = ca.getAll(keys), s = '', ok = true;
    for (i = 0; i < keys.length; i++) { if (got[keys[i]] == null) { ok = false; break; } s += got[keys[i]]; }
    if (ok) { try { return (MEMO[n] = JSON.parse(s)); } catch (e) { } }
  }
  var out = readSheet_(n);
  try {
    var js = JSON.stringify(out);
    if (js.length < 800000) {
      var put = {}, cnt = Math.ceil(js.length / 90000) || 1;
      for (var k = 0; k < cnt; k++) put['d_' + n + v + '_' + k] = js.substr(k * 90000, 90000);
      ca.putAll(put, 1800); ca.put(mk, String(cnt), 1800);
    }
  } catch (e) { }
  return (MEMO[n] = out);
}
// Baris baru diberi format teks ('@') lebih dulu agar jam tidak diubah otomatis menjadi tanggal oleh Google Sheets.
function append_(n, o) { appendMany_(n, [o]); }
function appendMany_(n, list) {
  if (!list.length) return;
  var h = SCHEMA[n], sh = sheet_(n), start = sh.getLastRow() + 1;
  var need = start + list.length - 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
  var rng = sh.getRange(start, 1, list.length, h.length);
  rng.setNumberFormat('@');
  rng.setValues(list.map(function (o) { return h.map(function (k) { return o[k] == null ? '' : o[k]; }); }));
  bump_(n);
}
function write_(n, o) { var h = SCHEMA[n]; sheet_(n).getRange(o._row, 1, 1, h.length).setValues([h.map(function (k) { return o[k] == null ? '' : o[k]; })]); bump_(n); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function withLock_(fn) { var l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }
function nextId_(key, prefix, pad) {
  var p = PropertiesService.getScriptProperties(), n = Number(p.getProperty(key) || 0) + 1;
  p.setProperty(key, String(n));
  return prefix + ('000000' + n).slice(-pad);
}
function txt_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max || 200); }
function need_(c, m) { if (!c) throw new Error(m); }
function audit_(u, aksi, detail) {
  try { append_('AuditLog', { Waktu: now_(), User: u ? u.Username : '-', Role: u ? u.Role : '-', Aksi: aksi, Detail: String(detail || '').slice(0, 500) }); } catch (e) { }
}

/* ============================ AUTH ============================ */
function hash_(pw, salt) {
  var pepper = PropertiesService.getScriptProperties().getProperty('PEPPER') || '';
  var h = salt + pepper + pw;
  for (var i = 0; i < 200; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  return h;
}
function checkPw_(pw) {
  need_(typeof pw === 'string' && pw.length >= 8 && pw.length <= 64 && /[A-Za-z]/.test(pw) && /\d/.test(pw), 'Password minimal 8 karakter dan mengandung huruf serta angka');
}
function findUser_(un) { un = String(un || '').toLowerCase(); return readAll_('Users').filter(function (u) { return u.Username === un; })[0]; }
function createUser_(p) {
  var un = String(p.username || '').trim().toLowerCase();
  need_(/^[a-z0-9._]{3,30}$/.test(un), 'Username 3-30 karakter (huruf kecil, angka, titik, underscore)');
  need_(!findUser_(un), 'Username sudah dipakai');
  checkPw_(p.password);
  var salt = Utilities.getUuid();
  append_('Users', {
    ID: nextId_('USEQ', 'U', 4), Username: un, Hash: hash_(p.password, salt), Salt: salt, Nama: txt_(p.nama, 60),
    Role: p.role, Tipe: p.tipe || '', CustomerIDs: p.customerIDs || '', Aktif: p.aktif || 'Y', Dibuat: now_(), LoginTerakhir: ''
  });
}
function pubUser_(u) { return { username: u.Username, nama: u.Nama, role: u.Role, tipe: u.Tipe, customerIDs: u.CustomerIDs ? u.CustomerIDs.split(',') : [] }; }
function auth_(t) {
  if (!t) throw new Error('SESSION');
  var c = CacheService.getScriptCache(), id = c.get('s_' + t);
  if (!id) throw new Error('SESSION');
  var u = readAll_('Users').filter(function (x) { return x.ID === id; })[0];
  if (!u || u.Aktif !== 'Y') throw new Error('SESSION');
  c.put('s_' + t, id, CFG.TTL);
  return u;
}
function tryAuth_(t) { try { return t ? auth_(t) : null; } catch (e) { return null; } }

function bundleFor_(u) {
  var b = {
    user: pubUser_(u),
    cfg: cfgOut_(),
    intDash: dashboard_({ filters: {} }, u),
    pubDash: dashboard_({ filters: {} }, null),
    inbox: chatInbox_({}, u)
  };
  if (u.Role === 'LEADER' || u.Role === 'ADMIN') b.fd = formData_({}, u);
  if (u.Role === 'ADMIN') {
    b.users = userList_();
    b.customers = masterList_({ type: 'Customers' });
    b.checkers = masterList_({ type: 'Checkers' });
    b.audit = auditList_();
  }
  return b;
}

function login_(p) {
  var un = String(p.username || '').trim().toLowerCase(), pw = String(p.password || '');
  need_(un && pw, 'Username dan password wajib diisi');
  var c = CacheService.getScriptCache(), k = 'f_' + un, f = Number(c.get(k) || 0);
  need_(f < CFG.MAX_FAIL, 'Terlalu banyak percobaan gagal. Coba lagi dalam 15 menit.');
  var u = findUser_(un);
  var ok = u && u.Aktif === 'Y' && hash_(pw, u.Salt) === u.Hash;
  if (!ok) { c.put(k, String(f + 1), CFG.LOCK_SEC); audit_({ Username: un, Role: '-' }, 'LOGIN_GAGAL', ''); throw new Error('Username atau password salah'); }
  c.remove(k);
  var t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  c.put('s_' + t, u.ID, CFG.TTL);
  u.LoginTerakhir = now_(); write_('Users', u);
  audit_(u, 'LOGIN', '');
  return { token: t, user: pubUser_(u), defaultPw: (un === CFG.DEFAULT_ADMIN.username && pw === adminPw_()), bundle: bundleFor_(u) };
}

/* ============================ API ROUTER ============================ */
function api(action, token, p) {
  try {
    p = p || {};
    var r = ROUTES[action];
    if (!r) throw new Error('Aksi tidak dikenal');
    var user = null;
    if (r.roles) {
      user = auth_(token);
      if (r.roles.indexOf(user.Role) < 0) { audit_(user, 'AKSES_DITOLAK', action); throw new Error('Akses ditolak untuk peran Anda'); }
    } else user = tryAuth_(token);
    return { ok: true, data: r.fn(p, user, token) };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}
var ALL = ['ADMIN', 'LEADER', 'PENGAWAS'];
var ROUTES = {
  login: { fn: function (p) { return login_(p); } },
  logout: { fn: function (p, u, t) { if (u) { audit_(u, 'LOGOUT', ''); CacheService.getScriptCache().remove('s_' + t); } return true; } },
  me: { roles: ALL, fn: function (p, u) { return bundleFor_(u); } },
  cfg: { fn: function () { return cfgOut_(); } },
  changePassword: { roles: ALL, fn: changePassword_ },
  dashboard: { fn: dashboard_ },
  chatStart: { fn: chatStart_ },
  chatSend: { fn: chatSend_ },
  chatPoll: { fn: chatPoll_ },
  formData: { roles: ['LEADER', 'ADMIN'], fn: formData_ },
  jobSave: { roles: ['LEADER', 'ADMIN'], fn: jobSave_ },
  jobDelete: { roles: ['LEADER', 'ADMIN'], fn: jobDelete_ },
  jobStatus: { roles: ['PENGAWAS', 'ADMIN'], fn: jobStatus_ },
  jobImport: { roles: ['LEADER', 'ADMIN'], fn: jobImport_ },
  jobArea: { roles: ['PENGAWAS', 'ADMIN'], fn: jobArea_ },
  chatDelete: { roles: ALL, fn: chatDelete_ },
  jobCtns: { roles: ALL, fn: jobCtns_ },
  ctnToggle: { roles: ['PENGAWAS', 'ADMIN'], fn: ctnToggle_ },
  chatInbox: { roles: ALL, fn: chatInbox_ },
  chatThread: { roles: ALL, fn: chatThread_ },
  chatReply: { roles: ALL, fn: chatReply_ },
  userList: { roles: ['ADMIN'], fn: userList_ },
  userSave: { roles: ['ADMIN'], fn: userSave_ },
  masterList: { roles: ['ADMIN'], fn: masterList_ },
  masterSave: { roles: ['ADMIN'], fn: masterSave_ },
  auditList: { roles: ['ADMIN'], fn: auditList_ }
};
function cfgOut_() {
  return {
    kategori: CFG.KATEGORI,
    ukuran: CFG.UKURAN,
    jenisKargo: CFG.JENIS_KARGO,
    status: CFG.STATUS,
    statusPengawas: CFG.STATUS_PENGAWAS,
    hasil: CFG.HASIL,
    area: CFG.AREA,
    customers: readAll_('Customers').filter(function (c) { return c.TampilPublik === 'Y' && c.Aktif === 'Y'; }).map(function (c) { return { id: c.ID, nama: c.Nama }; })
  };
}

function changePassword_(p, u) {
  need_(hash_(String(p.oldPassword || ''), u.Salt) === u.Hash, 'Password lama salah');
  checkPw_(p.newPassword);
  u.Salt = Utilities.getUuid(); u.Hash = hash_(p.newPassword, u.Salt); write_('Users', u);
  audit_(u, 'GANTI_PASSWORD', ''); return true;
}

/* ============================ DASHBOARD ============================ */
function fmt_(d, f) { return Utilities.formatDate(d, TZ, f); }
function urg_(j, nowS, today) {
  if (j.Status === 'Selesai' || !j.Closing) return '';
  if (j.Closing < nowS) return 'TERLAMBAT';
  if (j.Closing.substr(0, 10) === today) return 'URGENT';
  return '';
}
function inScope_(j, u, cm) {
  if (!u) { var c = cm[j.CustomerID]; return !!c && c.TampilPublik === 'Y'; }
  if (u.Role === 'ADMIN') return true;
  if (u.Role === 'LEADER') return j.LeaderUser === u.Username;
  // Pengawas melihat tugas miliknya ATAU tugas yang pengawasnya belum ditentukan
  return j.PengawasUser === u.Username || !j.PengawasUser;
}
function row_(j, internal) {
  var r = {
    id: j.ID, tgl: j.Tanggal, custId: j.CustomerID, cust: j.Customer, kat: j.Kategori,
    status: normStatus_(j.Status), no: j.NoContainer, uk: j.Ukuran,
    jmlArmada: j.JmlArmada || '1', jenisKargo: j.JenisKargo || '-', area: j.Area || '',
    close: j.Closing, leader: j.LeaderNama, pengawas: j.PengawasNama || '-',
    hasil: j.Hasil, urg: j._u, selesai: j.Selesai,
    nCtn: j._n == null ? 1 : j._n, nDone: j._d == null ? 0 : j._d
  };
  if (internal) {
    r.detail = j.Detail;
    r.checker = j.Checker || '-';
    r.catatan = j.Catatan;
    r.pengawasUser = j.PengawasUser || '';
    r.upd = j.Diupdate;
  }
  return r;
}
function uniqBy_(arr, kf, mf) { var s = {}, o = []; arr.forEach(function (x) { var k = kf(x); if (k && !s[k]) { s[k] = 1; o.push(mf(x)); } }); return o.sort(function (a, b) { return (a.nama || '').localeCompare(b.nama || ''); }); }

function dashboard_(p, user) {
  var cm = {}; readAll_('Customers').forEach(function (c) { cm[c.ID] = c; });
  var nowS = fmt_(new Date(), 'yyyy-MM-dd HH:mm'), today = nowS.substr(0, 10);
  var scoped = readAll_('Jobs').filter(function (j) { return inScope_(j, user, cm); });
  var cs = ctnStats_();
  scoped.forEach(function (j) {
    j._u = urg_(j, nowS, today);
    var st = cs[j.ID], legacy = splitNos_(j.NoContainer).length || 1;
    j._n = st ? st.n : legacy;
    j._d = st ? st.d : (normStatus_(j.Status) === 'Selesai' ? legacy : 0);
  });
  var opts = {
    customers: uniqBy_(scoped, function (j) { return j.CustomerID; }, function (j) { return { id: j.CustomerID, nama: j.Customer }; }),
    leaders: uniqBy_(scoped, function (j) { return j.LeaderNama; }, function (j) { return { nama: j.LeaderNama }; }),
    pengawas: uniqBy_(scoped, function (j) { return j.PengawasNama; }, function (j) { return { nama: j.PengawasNama }; })
  };
  var rows = scoped.slice().sort(function (a, b) { return a.Tanggal === b.Tanggal ? (a.Dibuat < b.Dibuat ? 1 : -1) : (a.Tanggal < b.Tanggal ? 1 : -1); });
  var ids = {}; rows.forEach(function (j) { ids[j.ID] = 1; });
  var hist = readAll_('History').filter(function (h) { return ids[h.JobID]; }).reverse().slice(0, 80)
    .map(function (h) { return { jobId: h.JobID, waktu: h.Waktu, no: h.NoContainer, cust: h.Customer, dari: normStatus_(h.Dari), ke: normStatus_(h.Ke), hasil: h.Hasil, cat: user ? h.Catatan : '', oleh: h.Oleh, peran: h.Peran }; });
  var out = {
    opts: opts,
    history: hist,
    now: nowS,
    rows: rows.slice(0, 1500).map(function (j) { return row_(j, !!user); })
  };
  if (user) {
    out.alerts = scoped.filter(function (j) { return j._u; }).sort(function (a, b) { return a._u === b._u ? (a.Closing > b.Closing ? 1 : -1) : (a._u === 'TERLAMBAT' ? -1 : 1); })
      .slice(0, 30).map(function (j) { return row_(j, true); });
  }
  return out;
}

/* ============================ PEKERJAAN ============================ */
function formData_(p, u) {
  var cust = readAll_('Customers').filter(function (c) { return c.Aktif === 'Y'; });
  if (u.Role === 'LEADER') { var a = u.CustomerIDs ? u.CustomerIDs.split(',') : []; cust = cust.filter(function (c) { return a.indexOf(c.ID) >= 0; }); }
  return {
    customers: cust.map(function (c) { return { id: c.ID, nama: c.Nama }; }),
    checkers: readAll_('Checkers').filter(function (c) { return c.Aktif === 'Y'; }).map(function (c) { return { id: c.ID, nama: c.Nama }; }),
    pengawas: readAll_('Users').filter(function (x) { return x.Role === 'PENGAWAS' && x.Aktif === 'Y'; }).map(function (x) { return { user: x.Username, nama: x.Nama }; })
  };
}
function hist_(j, dari, ke, hasil, cat, u) {
  append_('History', { Waktu: now_(), JobID: j.ID, NoContainer: j.NoContainer, Customer: j.Customer, Dari: dari, Ke: ke, Hasil: hasil || '', Catatan: cat || '', Oleh: u.Nama, Peran: u.Role });
}
function jobSave_(p, u) {
  return withLock_(function () {
    var edit = !!p.id, jobs = readAll_('Jobs'), old = null;
    if (edit) {
      old = jobs.filter(function (j) { return j.ID === p.id; })[0];
      need_(old, 'Pekerjaan tidak ditemukan');
      need_(u.Role === 'ADMIN' || old.LeaderUser === u.Username, 'Akses ditolak');
    } else need_(u.Role === 'LEADER' || u.Role === 'ADMIN', 'Hanya Leader yang dapat menambah pekerjaan');

    var tgl = txt_(p.tanggal, 10); need_(/^\d{4}-\d{2}-\d{2}$/.test(tgl), 'Tanggal tidak valid');
    var close = txt_(p.closing, 16).replace('T', ' '); need_(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(close), 'Closing date/time tidak valid');
    // Satu aktivitas boleh berisi banyak nomor container/plat (pisahkan dengan Enter, koma, titik koma, atau spasi)
    var nos = [], seenNo = {};
    txt_(p.noContainer, 6000).toUpperCase().split(/[\s,;]+/).forEach(function (t) {
      t = t.replace(/-/g, '');
      if (!t) return;
      need_(/^[A-Z0-9]{6,15}$/.test(t), 'Nomor container / plat "' + t + '" tidak valid (6-15 huruf/angka)');
      need_(!seenNo[t], 'Nomor ' + t + ' diinput dobel pada pekerjaan ini');
      seenNo[t] = 1; nos.push(t);
    });
    need_(nos.length > 0, 'Nomor container / plat wajib diisi');
    need_(nos.length <= 200, 'Maksimal 200 nomor container per aktivitas');
    var no = nos.join(', ');
    need_(CFG.KATEGORI.indexOf(p.kategori) >= 0, 'Kategori tidak valid');
    need_(CFG.UKURAN.indexOf(p.ukuran) >= 0, 'Ukuran container/armada tidak valid');

    var jmlArmada = nos.length > 1 ? String(nos.length) : txt_(p.jmlArmada, 20);
    need_(jmlArmada && /^[0-9]+(\s*[A-Za-z0-9\/]*)?$/.test(jmlArmada), 'Jumlah Kontainer / Armada wajib diisi (contoh: 1, 2 Unit)');
    var jenisKargo = txt_(p.jenisKargo, 80);
    need_(jenisKargo.length >= 2, 'Jenis Kargo wajib diisi / dipilih');

    var detail = txt_(p.detail, 500); need_(detail.length >= 3, 'Detail pekerjaan wajib diisi');

    var cust = readAll_('Customers').filter(function (c) { return c.ID === p.customerId; })[0];
    need_(cust, 'Customer tidak valid');
    var unchanged = old && old.CustomerID === cust.ID;
    if (!unchanged) {
      need_(cust.Aktif === 'Y', 'Customer tidak aktif');
      if (u.Role === 'LEADER') need_((u.CustomerIDs || '').split(',').indexOf(cust.ID) >= 0, 'Customer ini tidak ditetapkan untuk Anda');
    }

    // Checker bersifat opsional
    var chkNama = '';
    if (p.checkerId) {
      var chk = readAll_('Checkers').filter(function (c) { return c.ID === p.checkerId; })[0];
      need_(chk && (chk.Aktif === 'Y' || (old && old.Checker === chk.Nama)), 'Checker tidak valid');
      chkNama = chk.Nama;
    }

    // Pengawas bersifat opsional
    var pgUser = '', pgNama = '';
    if (p.pengawas) {
      var pg = readAll_('Users').filter(function (x) { return x.Role === 'PENGAWAS' && x.Username === p.pengawas; })[0];
      need_(pg && (pg.Aktif === 'Y' || (old && old.PengawasUser === pg.Username)), 'Pengawas tidak valid');
      pgUser = pg.Username;
      pgNama = pg.Nama;
    }

    // Status otomatis 'Belum Dikerjakan' untuk data baru, atau pertahankan status lama saat edit
    var statusJob = edit ? normStatus_(old.Status) : 'Belum Dikerjakan';

    if (statusJob !== 'Selesai') {
      var dupNo = '';
      jobs.forEach(function (j) {
        if (dupNo || j.ID === p.id || j.CustomerID !== cust.ID || normStatus_(j.Status) === 'Selesai') return;
        dupNo = overlap_(splitNos_(j.NoContainer), nos);
      });
      need_(!dupNo, 'Container ' + dupNo + ' masih outstanding pada customer ini (satu baris = satu container)');
    }

    var leader = old ? { user: old.LeaderUser, nama: old.LeaderNama } : { user: u.Username, nama: u.Nama };
    var j = old || { ID: nextId_('JSEQ', 'JOB-', 6), Dibuat: now_(), Hasil: '', Catatan: '', Selesai: '' };
    var prevStatus = old ? normStatus_(old.Status) : '-';
    var oldNoStr = old ? old.NoContainer : '';
    j.Tanggal = tgl; j.CustomerID = cust.ID; j.Customer = cust.Nama; j.Kategori = p.kategori; j.Status = statusJob; j.Detail = detail;
    j.Ukuran = p.ukuran; j.NoContainer = no; j.Closing = close; j.Checker = chkNama; j.PengawasUser = pgUser; j.PengawasNama = pgNama;
    j.JmlArmada = jmlArmada; j.JenisKargo = jenisKargo;
    j.LeaderUser = leader.user; j.LeaderNama = leader.nama; j.Diupdate = now_(); j.DiupdateOleh = u.Nama;

    if (edit) {
      // Data lama tanpa baris Containers: materialisasi dulu supaya status Done tidak hilang
      var oldNos = splitNos_(oldNoStr), oldDone = prevStatus === 'Selesai';
      if (!readAll_('Containers').some(function (c) { return c.JobID === j.ID; })) syncCtns_(j.ID, oldNos, oldDone);
      // Kontainer baru ditambahkan ke pekerjaan yang sudah Selesai -> kembali On Progress
      var added = nos.filter(function (n) { return oldNos.indexOf(n) < 0; });
      if (prevStatus === 'Selesai' && added.length) { j.Status = 'On Progress'; j.Hasil = ''; j.Selesai = ''; }
      write_('Jobs', j);
      if (j.Status !== prevStatus) hist_(j, prevStatus, j.Status, '', 'Kontainer ditambahkan, status kembali On Progress', u);
    } else append_('Jobs', j);
    syncCtns_(j.ID, nos, false);
    if (!edit) hist_(j, prevStatus, j.Status, '', 'Pekerjaan diinput oleh Leader', u);
    audit_(u, edit ? 'JOB_EDIT' : 'JOB_BUAT', j.ID + ' ' + j.NoContainer + ' ' + j.Customer);
    return { id: j.ID, intDash: dashboard_({}, u), pubDash: dashboard_({}, null) };
  });
}
function jobDelete_(p, u) {
  return withLock_(function () {
    var j = readAll_('Jobs').filter(function (x) { return x.ID === p.id; })[0];
    need_(j, 'Pekerjaan tidak ditemukan');
    need_(u.Role === 'ADMIN' || j.LeaderUser === u.Username, 'Akses ditolak');
    sheet_('Jobs').deleteRow(j._row); bump_('Jobs');
    syncCtns_(j.ID, [], false);
    hist_(j, normStatus_(j.Status), 'DIHAPUS', '', '', u);
    audit_(u, 'JOB_HAPUS', j.ID + ' ' + j.NoContainer);
    return { intDash: dashboard_({}, u), pubDash: dashboard_({}, null) };
  });
}
function jobStatus_(p, u) {
  return withLock_(function () {
    var j = readAll_('Jobs').filter(function (x) { return x.ID === p.id; })[0];
    need_(j, 'Pekerjaan tidak ditemukan');
    need_(u.Role === 'ADMIN' || !j.PengawasUser || j.PengawasUser === u.Username, 'Pekerjaan ini tidak ditugaskan kepada Anda');
    need_(CFG.STATUS_PENGAWAS.indexOf(p.status) >= 0, 'Status harus dipilih antara On Progress atau Selesai');
    var area = p.area ? String(p.area) : (j.Area || '');
    need_(area, 'Pilih area yang dikerjakan terlebih dahulu (Phase 2, Phase 4, atau Phase 8)');
    need_(CFG.AREA.indexOf(area) >= 0, 'Area tidak valid');
    var areaChanged = (j.Area || '') !== area;
    var hasil = p.hasil ? String(p.hasil) : '', cat = txt_(p.catatan, 300);
    need_(!hasil || CFG.HASIL.indexOf(hasil) >= 0, 'Hasil tidak valid');
    if (p.status === 'Selesai') need_(hasil, 'Hasil wajib diisi saat status Selesai');
    if (hasil === 'Lainnya' || hasil === 'Damage' || hasil === 'Overload') need_(cat, 'Isi catatan untuk hasil ' + hasil);
    if (p.status === 'Selesai') {
      var ctn0 = ctnsOf_(j), belum = ctn0.filter(function (c) { return !c.d; });
      if (ctn0.length > 1) need_(!belum.length, belum.length + ' kontainer belum Done (' + belum.slice(0, 3).map(function (c) { return c.no; }).join(', ') + (belum.length > 3 ? ', ...' : '') + '). Centang semua kontainer terlebih dahulu.');
      else if (belum.length) setCtns_(j, belum.map(function (c) { return c.no; }), true, u);
    }
    var dari = normStatus_(j.Status), hLama = j.Hasil;
    j.Status = p.status; j.Hasil = p.status === 'Selesai' ? hasil : '';
    j.Area = area;
    if (cat) j.Catatan = cat;
    // Jika sebelumnya pengawas belum ditentukan dan diupdate oleh Pengawas, otomatis catat pengawas tersebut
    if (!j.PengawasUser && u.Role === 'PENGAWAS') {
      j.PengawasUser = u.Username;
      j.PengawasNama = u.Nama;
    }
    j.Selesai = p.status === 'Selesai' ? (j.Selesai || now_()) : '';
    j.Diupdate = now_(); j.DiupdateOleh = u.Nama;
    write_('Jobs', j);
    if (dari !== j.Status || hLama !== j.Hasil || cat || areaChanged) hist_(j, dari, j.Status, j.Hasil, (areaChanged ? 'Area: ' + area : '') + (areaChanged && cat ? ' · ' : '') + cat, u);
    audit_(u, 'JOB_STATUS', j.ID + ' ' + dari + '->' + j.Status + ' ' + j.Hasil + ' [' + area + ']');
    return { intDash: dashboard_({}, u), pubDash: dashboard_({}, null) };
  });
}

/* ============================ KONTAINER PER AKTIVITAS ============================ */
function splitNos_(s) { return String(s == null ? '' : s).split(',').map(function (x) { return x.trim(); }).filter(Boolean); }
function overlap_(a, b) { for (var i = 0; i < a.length; i++) if (b.indexOf(a[i]) >= 0) return a[i]; return ''; }
function ctnStats_() {
  var m = {};
  readAll_('Containers').forEach(function (c) {
    var o = m[c.JobID] || (m[c.JobID] = { n: 0, d: 0 });
    o.n++; if (c.Done === 'Y') o.d++;
  });
  return m;
}
// Daftar kontainer sebuah pekerjaan. Data lama (tanpa baris di sheet Containers) diturunkan dari kolom NoContainer.
function ctnsOf_(j) {
  var rows = readAll_('Containers').filter(function (c) { return c.JobID === j.ID; });
  if (rows.length) return rows.map(function (c) { return { no: c.NoContainer, d: c.Done === 'Y', at: c.DoneAt, by: c.DoneBy }; });
  var selesai = normStatus_(j.Status) === 'Selesai';
  return splitNos_(j.NoContainer).map(function (n) { return { no: n, d: selesai, at: selesai ? j.Selesai : '', by: selesai ? j.DiupdateOleh : '' }; });
}
// Samakan baris sheet Containers dengan daftar nomor (tambah yang baru, hapus yang tidak ada lagi)
function syncCtns_(jobId, nos, doneDefault) {
  var rows = readAll_('Containers').filter(function (c) { return c.JobID === jobId; });
  var keep = {}, add = [];
  nos.forEach(function (n) { keep[n] = 1; });
  rows.filter(function (c) { return !keep[c.NoContainer]; }).sort(function (a, b) { return b._row - a._row; })
    .forEach(function (c) { sheet_('Containers').deleteRow(c._row); });
  if (rows.some(function (c) { return !keep[c.NoContainer]; })) bump_('Containers');
  var have = {}; rows.forEach(function (c) { have[c.NoContainer] = 1; });
  nos.forEach(function (n) {
    if (!have[n]) add.push({ JobID: jobId, NoContainer: n, Done: doneDefault ? 'Y' : 'N', DoneAt: doneDefault ? now_() : '', DoneBy: '' });
  });
  appendMany_('Containers', add);
}
function setCtns_(j, nos, done, u) {
  var rows = readAll_('Containers').filter(function (c) { return c.JobID === j.ID; });
  if (!rows.length) {   // materialisasi data lama
    syncCtns_(j.ID, splitNos_(j.NoContainer), normStatus_(j.Status) === 'Selesai');
    rows = readAll_('Containers').filter(function (c) { return c.JobID === j.ID; });
  }
  var set = {}; nos.forEach(function (n) { set[n] = 1; });
  rows.forEach(function (c) {
    if (set[c.NoContainer] && (c.Done === 'Y') !== done) {
      c.Done = done ? 'Y' : 'N'; c.DoneAt = done ? now_() : ''; c.DoneBy = done ? u.Nama : '';
      write_('Containers', c);
    }
  });
}
function jobCtns_(p, u) {
  var j = readAll_('Jobs').filter(function (x) { return x.ID === p.id; })[0];
  need_(j, 'Pekerjaan tidak ditemukan');
  need_(inScope_(j, u, {}), 'Akses ditolak');
  return ctnOut_(j);
}
function ctnOut_(j) {
  var list = ctnsOf_(j);
  return {
    id: j.ID, cust: j.Customer, kat: j.Kategori, uk: j.Ukuran, status: normStatus_(j.Status), close: j.Closing, area: j.Area || '',
    n: list.length, done: list.filter(function (c) { return c.d; }).length, ctns: list
  };
}
// Pengawas memilih area pengerjaan (Phase 2/4/8). Jika pekerjaan belum dimulai, otomatis menjadi On Progress.
function jobArea_(p, u) {
  return withLock_(function () {
    var j = readAll_('Jobs').filter(function (x) { return x.ID === p.id; })[0];
    need_(j, 'Pekerjaan tidak ditemukan');
    need_(u.Role === 'ADMIN' || !j.PengawasUser || j.PengawasUser === u.Username, 'Pekerjaan ini tidak ditugaskan kepada Anda');
    var area = String(p.area || '');
    need_(CFG.AREA.indexOf(area) >= 0, 'Pilih area: ' + CFG.AREA.join(', '));
    var dari = normStatus_(j.Status), lama = j.Area || '';
    if (lama !== area || dari === 'Belum Dikerjakan') {
      j.Area = area;
      if (!j.PengawasUser && u.Role === 'PENGAWAS') { j.PengawasUser = u.Username; j.PengawasNama = u.Nama; }
      if (dari === 'Belum Dikerjakan') j.Status = 'On Progress';
      j.Diupdate = now_(); j.DiupdateOleh = u.Nama;
      write_('Jobs', j);
      hist_(j, dari, normStatus_(j.Status), '', lama ? 'Area diubah: ' + lama + ' → ' + area : 'Mulai dikerjakan di ' + area, u);
      audit_(u, 'JOB_AREA', j.ID + ' ' + (lama || '-') + '->' + area);
    }
    var out = ctnOut_(j);
    out.intDash = dashboard_({}, u); out.pubDash = dashboard_({}, null);
    return out;
  });
}
function ctnToggle_(p, u) {
  return withLock_(function () {
    var j = readAll_('Jobs').filter(function (x) { return x.ID === p.id; })[0];
    need_(j, 'Pekerjaan tidak ditemukan');
    need_(u.Role === 'ADMIN' || !j.PengawasUser || j.PengawasUser === u.Username, 'Pekerjaan ini tidak ditugaskan kepada Anda');
    var req = (Array.isArray(p.nos) ? p.nos : [p.no]).map(function (x) { return String(x == null ? '' : x).trim(); }).filter(Boolean);
    need_(req.length > 0, 'Pilih kontainer terlebih dahulu');
    var done = p.done === true || p.done === 'true';
    if (done) need_(j.Area, 'Pilih area yang dikerjakan (Phase 2 / 4 / 8) sebelum mencentang kontainer');
    var cur = ctnsOf_(j), valid = {}; cur.forEach(function (c) { valid[c.no] = 1; });
    req.forEach(function (n) { need_(valid[n], 'Kontainer ' + n + ' tidak ada pada pekerjaan ini'); });
    var changed = req.filter(function (n) { return cur.filter(function (c) { return c.no === n; })[0].d !== done; });
    if (changed.length) {
      setCtns_(j, changed, done, u);
      var dari = normStatus_(j.Status);
      if (!j.PengawasUser && u.Role === 'PENGAWAS') { j.PengawasUser = u.Username; j.PengawasNama = u.Nama; }
      if (done && dari === 'Belum Dikerjakan') j.Status = 'On Progress';
      if (!done && dari === 'Selesai') { j.Status = 'On Progress'; j.Hasil = ''; j.Selesai = ''; }
      j.Diupdate = now_(); j.DiupdateOleh = u.Nama;
      write_('Jobs', j);
      append_('History', { Waktu: now_(), JobID: j.ID, NoContainer: changed.slice(0, 3).join(', ') + (changed.length > 3 ? ' +' + (changed.length - 3) : ''), Customer: j.Customer,
        Dari: done ? 'Belum Done' : 'Done', Ke: done ? 'Done' : 'Belum Done', Hasil: '', Catatan: '', Oleh: u.Nama, Peran: u.Role });
      if (dari !== normStatus_(j.Status)) hist_(j, dari, normStatus_(j.Status), '', 'Status otomatis menyesuaikan centang kontainer', u);
      audit_(u, 'CTN_' + (done ? 'DONE' : 'UNDONE'), j.ID + ' ' + changed.join(','));
    }
    var out = ctnOut_(j);
    out.intDash = dashboard_({}, u); out.pubDash = dashboard_({}, null);
    return out;
  });
}

/* ============================ INPUT EXCEL ============================ */
function normKey_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function pickList_(list, v) { var k = normKey_(v); if (!k) return ''; for (var i = 0; i < list.length; i++) if (normKey_(list[i]) === k) return list[i]; return ''; }
function validDate_(s) { var a = s.split('-').map(Number), d = new Date(Date.UTC(a[0], a[1] - 1, a[2])); return d.getUTCFullYear() === a[0] && d.getUTCMonth() === a[1] - 1 && d.getUTCDate() === a[2]; }

/**
 * Impor pekerjaan dari Excel. Satu baris Excel = satu nomor container/plat.
 * Baris dengan data yang sama (tanggal, customer, kategori, kargo, jumlah, ukuran, closing, checker, pengawas, detail)
 * digabung menjadi SATU aktivitas berisi beberapa kontainer. p.dry = true hanya memeriksa (pratinjau).
 */
function jobImport_(p, u) {
  var rows = p.rows;
  need_(Array.isArray(rows) && rows.length > 0, 'Tidak ada baris data pada file Excel');
  need_(rows.length <= 500, 'Maksimal 500 baris per unggahan (file berisi ' + rows.length + ' baris)');
  return withLock_(function () {
    var fd = formData_({}, u), errs = [], items = [];
    var chkAll = readAll_('Checkers').filter(function (c) { return c.Aktif === 'Y'; });
    var pgAll = readAll_('Users').filter(function (x) { return x.Role === 'PENGAWAS' && x.Aktif === 'Y'; });
    var err = function (r, m) { errs.push({ row: r, msg: m }); };

    rows.forEach(function (x) {
      var r = Number(x.r) || 0, e0 = errs.length;
      var tgl = txt_(x.tanggal, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(tgl) || !validDate_(tgl)) err(r, 'Tanggal tidak valid (gunakan format dd/mm/yyyy)');
      var cn = txt_(x.customer, 80).toLowerCase();
      var cust = fd.customers.filter(function (c) { return c.nama.toLowerCase() === cn || c.id === txt_(x.customer, 20); })[0];
      if (!cust) err(r, cn ? 'Customer "' + txt_(x.customer, 80) + '" tidak ditemukan atau tidak ditetapkan untuk Anda' : 'Customer wajib diisi');
      var kat = pickList_(CFG.KATEGORI, x.kategori); if (!kat) err(r, 'Kategori harus salah satu: ' + CFG.KATEGORI.join(', '));
      var kargo = pickList_(CFG.JENIS_KARGO, x.jenisKargo); if (!kargo) err(r, 'Jenis Kargo harus salah satu: ' + CFG.JENIS_KARGO.join(', '));
      var jml = txt_(x.jmlArmada, 20);
      if (!jml || !/^[0-9]+(\s*[A-Za-z0-9\/]*)?$/.test(jml) || parseInt(jml, 10) < 1) err(r, 'Jumlah Kontainer / Armada wajib diisi (contoh: 8)');
      var uk = pickList_(CFG.UKURAN, x.ukuran); if (!uk) err(r, 'Ukuran harus salah satu: ' + CFG.UKURAN.join(', '));
      var no = txt_(x.noContainer, 30).toUpperCase().replace(/[\s-]/g, '');
      if (!/^[A-Z0-9]{6,15}$/.test(no)) err(r, 'Nomor container / plat tidak valid (6-15 huruf/angka)');
      var close = txt_(x.closing, 16).replace('T', ' ');
      if (!/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(close)) err(r, 'Closing Date/Time tidak valid (gunakan dd/mm/yyyy hh:mm)');
      var ck = '', cks = txt_(x.checker, 80).toLowerCase();
      if (cks) { var chk = chkAll.filter(function (c) { return c.Nama.toLowerCase() === cks; })[0]; if (chk) ck = chk.Nama; else err(r, 'Checker "' + txt_(x.checker, 80) + '" tidak ditemukan'); }
      var pgUser = '', pgNama = '', pgs = txt_(x.pengawas, 80).toLowerCase();
      if (pgs) { var pg = pgAll.filter(function (q) { return q.Nama.toLowerCase() === pgs || q.Username === pgs; })[0]; if (pg) { pgUser = pg.Username; pgNama = pg.Nama; } else err(r, 'Pengawas "' + txt_(x.pengawas, 80) + '" tidak ditemukan'); }
      var detail = txt_(x.detail, 500); if (detail.length < 3) err(r, 'Detail Pekerjaan wajib diisi');
      if (errs.length === e0) items.push({ r: r, tgl: tgl, cust: cust, kat: kat, kargo: kargo, jml: jml, uk: uk, no: no, close: close, ck: ck, pgUser: pgUser, pgNama: pgNama, detail: detail });
    });

    // Gabungkan baris yang datanya sama menjadi satu aktivitas
    var groups = [], gmap = {};
    items.forEach(function (it) {
      var key = [it.tgl, it.cust.id, it.kat, it.kargo, it.jml, it.uk, it.close, it.ck, it.pgUser, it.detail].join('\u0001');
      var g = gmap[key]; if (!g) { g = gmap[key] = { first: it, rows: [], nos: [] }; groups.push(g); }
      g.rows.push(it.r); g.nos.push(it.no);
    });

    var seen = {}, out = {};
    groups.forEach(function (g) {
      var it = g.first, want = parseInt(it.jml, 10);
      g.nos.forEach(function (n, i) {
        var k = it.cust.id + '|' + n;
        if (seen[k]) err(g.rows[i], 'Nomor ' + n + ' tercantum dobel (juga di baris ' + seen[k] + ')'); else seen[k] = g.rows[i];
      });
      if (want !== g.nos.length) err(g.rows[0], 'Jumlah Kontainer / Armada = ' + it.jml + ' tetapi ada ' + g.nos.length + ' baris dengan data yang sama (' + it.cust.nama + '). Samakan jumlahnya, atau bedakan Detail Pekerjaan bila ini aktivitas terpisah.');
    });

    var outstanding = {};
    readAll_('Jobs').forEach(function (j) {
      if (normStatus_(j.Status) !== 'Selesai') splitNos_(j.NoContainer).forEach(function (n) { outstanding[j.CustomerID + '|' + n] = j.ID; });
    });
    groups.forEach(function (g) {
      g.nos.forEach(function (n, i) {
        var o = outstanding[g.first.cust.id + '|' + n];
        if (o) err(g.rows[i], 'Container ' + n + ' masih outstanding pada customer ini (' + o + ')');
      });
    });

    errs.sort(function (a, b) { return a.row - b.row; });
    out = {
      ok: errs.length === 0, dry: !!p.dry, saved: false, errCount: errs.length, errors: errs.slice(0, 100),
      nRows: rows.length, nJobs: groups.length, nCtn: items.length,
      groups: groups.slice(0, 200).map(function (g) {
        var it = g.first;
        return { tgl: it.tgl, cust: it.cust.nama, kat: it.kat, kargo: it.kargo, uk: it.uk, jml: it.jml, close: it.close, checker: it.ck || '-', pengawas: it.pgNama || '-', detail: it.detail, nos: g.nos };
      })
    };
    if (errs.length || p.dry) return out;

    var t = now_(), jobList = [], ctnList = [], histList = [];
    groups.forEach(function (g) {
      var it = g.first, id = nextId_('JSEQ', 'JOB-', 6);
      jobList.push({ ID: id, Tanggal: it.tgl, CustomerID: it.cust.id, Customer: it.cust.nama, Kategori: it.kat, Status: 'Belum Dikerjakan', Detail: it.detail,
        Ukuran: it.uk, NoContainer: g.nos.join(', '), Closing: it.close, Checker: it.ck, PengawasUser: it.pgUser, PengawasNama: it.pgNama,
        LeaderUser: u.Username, LeaderNama: u.Nama, Hasil: '', Catatan: '', Dibuat: t, Diupdate: t, DiupdateOleh: u.Nama, Selesai: '',
        JmlArmada: it.jml, JenisKargo: it.kargo });
      g.nos.forEach(function (n) { ctnList.push({ JobID: id, NoContainer: n, Done: 'N', DoneAt: '', DoneBy: '' }); });
      histList.push({ Waktu: t, JobID: id, NoContainer: g.nos.slice(0, 3).join(', ') + (g.nos.length > 3 ? ' +' + (g.nos.length - 3) : ''), Customer: it.cust.nama,
        Dari: '-', Ke: 'Belum Dikerjakan', Hasil: '', Catatan: 'Pekerjaan diinput via Excel (' + g.nos.length + ' kontainer)', Oleh: u.Nama, Peran: u.Role });
    });
    appendMany_('Jobs', jobList); appendMany_('Containers', ctnList); appendMany_('History', histList);
    audit_(u, 'JOB_IMPORT', jobList.length + ' aktivitas / ' + ctnList.length + ' kontainer');
    out.saved = true; out.intDash = dashboard_({}, u); out.pubDash = dashboard_({}, null);
    return out;
  });
}

/* ============================ LIVE CHAT CS (TANPA TIKET) ============================ */
function namesOf_(t) {
  var um = {}; readAll_('Users').forEach(function (u) { um[u.Username] = u.Nama; });
  var l = t.LeaderUser ? t.LeaderUser.split(',').filter(Boolean).map(function (x) { return um[x] || x; }) : [];
  var p = t.PengawasUser ? t.PengawasUser.split(',').filter(Boolean).map(function (x) { return um[x] || x; }) : [];
  return { leader: l, pengawas: p };
}
function rate_(key, max, sec) {
  var ca = CacheService.getScriptCache(), n = Number(ca.get(key) || 0);
  need_(n < max, 'Terlalu banyak pesan. Coba lagi sebentar lagi.'); ca.put(key, String(n + 1), sec);
}
function msgOut_(m) { return { dari: m.Dari, pengirim: m.Pengirim, pesan: m.Pesan, waktu: m.Waktu }; }
function chatFind_(id) {
  id = String(id || '');
  need_(/^[a-f0-9]{32}$/.test(id), 'CHAT_NOT_FOUND');
  var c = readAll_('Chats').filter(function (x) { return x.ChatID === id; })[0];
  need_(c, 'CHAT_NOT_FOUND'); return c;
}
function chatAccess_(c, u) {
  return u.Role === 'ADMIN' || (',' + c.LeaderUser + ',').indexOf(',' + u.Username + ',') >= 0 || (',' + c.PengawasUser + ',').indexOf(',' + u.Username + ',') >= 0;
}
function chatStart_(p) {
  var nama = txt_(p.nama, 60), pesan = txt_(p.pesan, 1000);
  need_(nama.length >= 2, 'Nama minimal 2 karakter'); need_(pesan.length >= 1, 'Pesan terlalu pendek');
  rate_('rl_all', 80, 600); rate_('rl_n_' + nama.toLowerCase(), 10, 600);
  var no = txt_(p.noContainer, 30).toUpperCase().replace(/[\s-]/g, '');
  if (no) need_(/^[A-Z0-9]{6,15}$/.test(no), 'Format nomor container tidak valid');
  var cm = {}; readAll_('Customers').forEach(function (x) { cm[x.ID] = x; });
  var cust = (p.customerId && cm[p.customerId] && cm[p.customerId].TampilPublik === 'Y') ? cm[p.customerId] : null;
  
  var jobs = readAll_('Jobs'), lead = {}, peng = {}, match = false;
  if (no) {
    jobs.filter(function (j) { return splitNos_(j.NoContainer).indexOf(no) >= 0; }).sort(function (a, b) { return a.Dibuat < b.Dibuat ? 1 : -1; })
      .slice(0, 1).forEach(function (j) { match = true; if (j.LeaderUser) lead[j.LeaderUser] = 1; if (j.PengawasUser) peng[j.PengawasUser] = 1; cust = cm[j.CustomerID] || cust; });
  }
  if (!match && cust) {
    readAll_('Users').forEach(function (x) { 
      if (x.Role === 'LEADER' && x.Aktif === 'Y' && (',' + x.CustomerIDs + ',').indexOf(',' + cust.ID + ',') >= 0) lead[x.Username] = 1; 
      if (x.Role === 'PENGAWAS' && x.Aktif === 'Y') peng[x.Username] = 1; // Otomatis teruskan ke semua pengawas aktif agar terpantau
    });
  }
  // Jika belum ada mapping spesifik, libatkan semua Leader & Pengawas aktif
  if (Object.keys(lead).length === 0 && Object.keys(peng).length === 0) {
    readAll_('Users').forEach(function (x) {
      if (x.Aktif === 'Y') {
        if (x.Role === 'LEADER') lead[x.Username] = 1;
        if (x.Role === 'PENGAWAS') peng[x.Username] = 1;
      }
    });
  }

  return withLock_(function () {
    var id = Utilities.getUuid().replace(/-/g, '');
    var c = { ChatID: id, Waktu: now_(), Nama: nama, Customer: cust ? cust.Nama : '', NoContainer: no, Status: 'UNREAD',
      LeaderUser: Object.keys(lead).join(','), PengawasUser: Object.keys(peng).join(','), Diupdate: now_(), DiupdateOleh: '' };
    append_('Chats', c);
    append_('ChatMsgs', { ChatID: id, Waktu: now_(), Dari: 'VISITOR', Pengirim: nama, Pesan: pesan });
    var n = namesOf_(c), fw = n.leader.concat(n.pengawas);
    append_('ChatMsgs', { ChatID: id, Waktu: now_(), Dari: 'BOT', Pengirim: 'CS Crossdocking',
      Pesan: 'Terima kasih, Kak ' + nama + '! Pesan Anda telah diteruskan ke tim Leader & Pengawas (' + (fw.length ? fw.join(', ') : 'Tim Operasional') + '). Mohon tunggu sebentar, balasan akan segera muncul di sini.' });
    audit_({ Username: 'publik', Role: 'PUBLIK' }, 'CHAT_MULAI', nama + (c.Customer ? ' / ' + c.Customer : ''));
    return { chatId: id };
  });
}
function chatSend_(p) {
  var c = chatFind_(p.chatId), pesan = txt_(p.pesan, 1000);
  need_(pesan.length >= 1, 'Pesan kosong');
  rate_('rl_c_' + c.ChatID, 40, 600);
  return withLock_(function () {
    append_('ChatMsgs', { ChatID: c.ChatID, Waktu: now_(), Dari: 'VISITOR', Pengirim: c.Nama, Pesan: pesan });
    c.Status = 'UNREAD'; c.FollowUpAt = '';   // pengunjung membalas -> batalkan pesan penutup otomatis
    c.Diupdate = now_(); write_('Chats', c); return true;
  });
}
// Waktu "yyyy-MM-dd HH:mm:ss" (WIB) -> epoch ms
function ts_(s) { return s ? new Date(String(s).replace(' ', 'T') + '+07:00').getTime() : 0; }
// Kirim pesan penutup otomatis bila pengunjung tidak membalas dalam 5 menit setelah dijawab petugas
function chatSweep_(list) {
  var t = Date.now();
  var due = list.filter(function (c) { return c.FollowUpAt && t - ts_(c.FollowUpAt) >= CFG.FOLLOWUP_MS; });
  if (!due.length) return;
  withLock_(function () {
    delete MEMO.Chats; delete MEMO.ChatMsgs;
    var fresh = readAll_('Chats');
    due.forEach(function (d) {
      var c = fresh.filter(function (x) { return x.ChatID === d.ChatID; })[0];
      if (!c || !c.FollowUpAt || Date.now() - ts_(c.FollowUpAt) < CFG.FOLLOWUP_MS) return;
      append_('ChatMsgs', { ChatID: c.ChatID, Waktu: now_(), Dari: 'BOT', Pengirim: 'CS Crossdocking', Pesan: 'Baik, terima kasih telah menghubungi layanan Crossdocking. Semoga informasi yang kami berikan bermanfaat. Selamat beraktivitas kembali' });
      c.FollowUpAt = ''; write_('Chats', c);
    });
  });
}
function chatPoll_(p) {
  var c = chatFind_(p.chatId), after = Number(p.after) || 0;
  chatSweep_([c]);
  var msgs = readAll_('ChatMsgs').filter(function (m) { return m.ChatID === c.ChatID; });
  var n = namesOf_(c);
  return { n: msgs.length, msgs: msgs.slice(after).map(msgOut_), diteruskan: n.leader.concat(n.pengawas) };
}
function chatInbox_(p, u) {
  chatSweep_(readAll_('Chats').filter(function (c) { return chatAccess_(c, u); }));
  var all = readAll_('ChatMsgs'), last = {}, cnt = {};
  all.forEach(function (m) { last[m.ChatID] = m; cnt[m.ChatID] = (cnt[m.ChatID] || 0) + 1; });
  return readAll_('Chats').filter(function (c) { return chatAccess_(c, u); })
    .sort(function (a, b) { return a.Diupdate < b.Diupdate ? 1 : -1; }).slice(0, 200).map(function (c) {
      var n = namesOf_(c), m = last[c.ChatID] || {};
      var unread = (c.Status === 'UNREAD' || c.Status === 'Baru') || m.Dari === 'VISITOR';
      return { id: c.ChatID, waktu: c.Waktu, nama: c.Nama, cust: c.Customer, container: c.NoContainer, unread: unread, diupdate: c.Diupdate,
        diteruskan: n.leader.concat(n.pengawas).join(', '), last: m.Pesan || '', lastDari: m.Dari || '', n: cnt[c.ChatID] || 0 };
    });
}
function chatThread_(p, u) {
  var c = chatFind_(p.id); need_(chatAccess_(c, u), 'Chat ini tidak diteruskan kepada Anda');
  chatSweep_([c]);
  if (c.Status === 'UNREAD' || c.Status === 'Baru') { c.Status = 'READ'; write_('Chats', c); }
  return { msgs: readAll_('ChatMsgs').filter(function (m) { return m.ChatID === c.ChatID; }).map(msgOut_) };
}
function chatReply_(p, u) {
  return withLock_(function () {
    var c = chatFind_(p.id); need_(chatAccess_(c, u), 'Chat ini tidak diteruskan kepada Anda');
    var rep = txt_(p.pesan, 1000);
    need_(rep.length >= 1, 'Tulis pesan balasan terlebih dahulu');
    append_('ChatMsgs', { ChatID: c.ChatID, Waktu: now_(), Dari: 'STAFF', Pengirim: u.Nama + ' (' + (u.Role === 'ADMIN' ? 'CS Admin' : u.Role === 'LEADER' ? 'Leader' : 'Pengawas') + ')', Pesan: rep });
    // Penutup otomatis: tanya apakah ada pertanyaan lain, lalu mulai hitung mundur 5 menit
    if (!c.FollowUpAt) {
      append_('ChatMsgs', { ChatID: c.ChatID, Waktu: now_(), Dari: 'BOT', Pengirim: 'CS Crossdocking',
        Pesan: 'Terima kasih, pertanyaan Anda telah kami jawab. Apakah masih ada hal lain yang dapat kami bantu?' });
    }
    c.FollowUpAt = now_();
    c.Status = 'READ'; c.Diupdate = now_(); c.DiupdateOleh = u.Nama; write_('Chats', c);
    audit_(u, 'CHAT_BALAS', c.Nama);
    return true;
  });
}
// Hapus room chat (beserta seluruh pesannya) agar inbox tidak menumpuk
function chatDelete_(p, u) {
  return withLock_(function () {
    var c = chatFind_(p.id); need_(chatAccess_(c, u), 'Chat ini tidak diteruskan kepada Anda');
    var rows = readAll_('ChatMsgs').filter(function (m) { return m.ChatID === c.ChatID; }).map(function (m) { return m._row; }).sort(function (a, b) { return b - a; });
    var sh = sheet_('ChatMsgs');
    rows.forEach(function (r) { sh.deleteRow(r); });
    bump_('ChatMsgs');
    sheet_('Chats').deleteRow(c._row); bump_('Chats');
    audit_(u, 'CHAT_HAPUS', c.Nama + (c.Customer ? ' / ' + c.Customer : '') + ' (' + rows.length + ' pesan)');
    return chatInbox_({}, u);
  });
}

/* ============================ ADMIN ============================ */
function userList_() {
  var cm = {}; readAll_('Customers').forEach(function (c) { cm[c.ID] = c.Nama; });
  return readAll_('Users').map(function (u) {
    var ids = u.CustomerIDs ? u.CustomerIDs.split(',').filter(Boolean) : [];
    return { id: u.ID, username: u.Username, nama: u.Nama, role: u.Role, tipe: u.Tipe, customerIDs: ids, customers: ids.map(function (i) { return cm[i] || i; }).join(', '), aktif: u.Aktif, login: u.LoginTerakhir };
  });
}
function userSave_(p, admin) {
  return withLock_(function () {
    var nama = txt_(p.nama, 60); need_(nama.length >= 2, 'Nama wajib diisi');
    need_(CFG.ROLES.indexOf(p.role) >= 0, 'Role tidak valid');
    var ids = [], tipe = '';
    if (p.role === 'LEADER') {
      var valid = readAll_('Customers').map(function (c) { return c.ID; });
      ids = (p.customerIDs || []).filter(function (i) { return valid.indexOf(i) >= 0; });
      tipe = ['Project', 'Non Project'].indexOf(p.tipe) >= 0 ? p.tipe : 'Project';
    }
    var aktif = p.aktif === 'N' ? 'N' : 'Y';
    if (!p.id) {
      createUser_({ username: p.username, nama: nama, role: p.role, tipe: tipe, customerIDs: ids.join(','), password: p.password, aktif: aktif });
      audit_(admin, 'USER_BUAT', p.username + ' (' + p.role + ')');
      return { users: userList_(), fd: formData_({}, admin) };
    }
    var all = readAll_('Users'), u = all.filter(function (x) { return x.ID === p.id; })[0];
    need_(u, 'Pengguna tidak ditemukan');
    if (u.ID === admin.ID) need_(p.role === 'ADMIN' && aktif === 'Y', 'Anda tidak dapat menurunkan/menonaktifkan akun sendiri');
    var wasAdmin = u.Role === 'ADMIN' && u.Aktif === 'Y';
    if (wasAdmin && (p.role !== 'ADMIN' || aktif !== 'Y'))
      need_(all.filter(function (x) { return x.Role === 'ADMIN' && x.Aktif === 'Y' && x.ID !== u.ID; }).length > 0, 'Harus ada minimal satu Admin aktif');
    u.Nama = nama; u.Role = p.role; u.Tipe = tipe; u.CustomerIDs = ids.join(','); u.Aktif = aktif;
    if (p.password) { checkPw_(p.password); u.Salt = Utilities.getUuid(); u.Hash = hash_(p.password, u.Salt); }
    write_('Users', u);
    if (aktif === 'N') CacheService.getScriptCache().remove('f_' + u.Username);
    audit_(admin, 'USER_UBAH', u.Username + ' role=' + u.Role + ' aktif=' + aktif + (p.password ? ' +reset_pw' : ''));
    return { users: userList_(), fd: formData_({}, admin) };
  });
}
function saveMaster_(sheetName, p, prefix) {
  var nama = txt_(p.nama, 80); need_(nama.length >= 2, 'Nama wajib diisi');
  var list = readAll_(sheetName);
  var dup = list.filter(function (x) { return x.Nama.toLowerCase() === nama.toLowerCase() && x.ID !== p.id; })[0];
  need_(!dup, 'Nama sudah ada');
  var o = p.id ? list.filter(function (x) { return x.ID === p.id; })[0] : null;
  if (p.id) need_(o, 'Data tidak ditemukan');
  o = o || { ID: nextId_(prefix + 'SEQ', prefix, 4) };
  o.Nama = nama; o.Aktif = p.aktif === 'N' ? 'N' : 'Y';
  if (sheetName === 'Customers') o.TampilPublik = p.tampilPublik === 'Y' ? 'Y' : 'N';
  if (p.id) write_(sheetName, o); else append_(sheetName, o);
  return o.ID;
}
function masterList_(p) {
  need_(p.type === 'Customers' || p.type === 'Checkers', 'Tipe tidak valid');
  return readAll_(p.type).map(function (x) { return { id: x.ID, nama: x.Nama, aktif: x.Aktif, tampilPublik: x.TampilPublik || '' }; });
}
function masterSave_(p, admin) {
  need_(p.type === 'Customers' || p.type === 'Checkers', 'Tipe tidak valid');
  return withLock_(function () {
    var id = saveMaster_(p.type, p, p.type === 'Customers' ? 'C' : 'K');
    audit_(admin, 'MASTER_' + p.type.toUpperCase(), id + ' ' + p.nama);
    return { id: id, list: masterList_(p), cfg: cfgOut_(), fd: formData_({}, admin) };
  });
}
function auditList_() {
  return readAll_('AuditLog').reverse().slice(0, 300).map(function (a) { return { waktu: a.Waktu, user: a.User, role: a.Role, aksi: a.Aksi, detail: a.Detail }; });
}