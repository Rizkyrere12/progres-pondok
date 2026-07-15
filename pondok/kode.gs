/**
 * ================================================================
 * SISTEM MANAJEMEN PONDOK — BACKEND (Google Apps Script)
 * ================================================================
 * File ini adalah API. Semua request dari HTML (login.html, index.html,
 * dst) dikirim ke sini lewat POST, dengan field "action" untuk menentukan
 * operasi apa yang dijalankan.
 *
 * CARA SETUP (lihat juga README.md):
 * 1. Buat Google Sheet baru (kosong).
 * 2. Buka Extensions > Apps Script, hapus isi default, tempel file ini.
 * 3. Di dropdown fungsi (atas, sebelah tombol Run), pilih "setupSheets"
 *    lalu klik Run sekali. Ini otomatis membuat 4 sheet + header kolom.
 * 4. Isi FONNTE_TOKEN di bawah kalau mau fitur kirim WA aktif (opsional).
 * 5. Deploy > New deployment > pilih tipe "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy "Web app URL" hasil deploy, dipakai di semua file HTML sebagai
 *    API_URL.
 * ================================================================
 */

const SHEET_SANTRI = 'DataSantri';
const SHEET_SETORAN = 'RiwayatSetoran';
const SHEET_USERS = 'Users';
const SHEET_KEUANGAN = 'TransaksiKeuangan';
const SHEET_ABSENSI = 'Absensi'; //FITUR BARU — sheet terpisah untuk modul absen (menggantikan pencatatan lama di RiwayatSetoran)
const FOLDER_UPLOAD_NAME = 'Upload_Pondok';

// Isi token Fonnte di sini kalau mau fitur WA aktif. Daftar di https://fonnte.com
const FONNTE_TOKEN = 'GANTI_DENGAN_TOKEN_FONNTE_KAMU';

// Role yang wajib verifikasi OTP WA setelah PIN benar (staff dengan akses sensitif)
const ROLE_WAJIB_OTP = ['SuperAdmin', 'Admin'];

/* ================= SETUP (dijalankan manual sekali) ================= */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schema = {
    [SHEET_SANTRI]: ['id', 'nama', 'kelas', 'halaqoh', 'no_wali', 'target', 'tercapai', 'target_kitab', 'tercapai_kitab', 'target_tahsin', 'tercapai_tahsin', 'poin', 'status', 'foto_url'],
    [SHEET_SETORAN]: ['id_santri', 'tanggal', 'jenis', 'jumlah', 'nilai', 'catatan', 'bukti_url', 'ustadz'],
    [SHEET_USERS]: ['username', 'pin', 'role', 'halaqoh', 'no_wa'],
    [SHEET_KEUANGAN]: ['tanggal', 'jenis', 'keterangan', 'jumlah', 'bukti']
  };
  Object.keys(schema).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const headers = schema[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  });
  // Hapus "Sheet1" default kalau masih ada dan kosong
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0) ss.deleteSheet(defaultSheet);

  // Tambahkan 1 akun SuperAdmin contoh supaya bisa langsung login pertama kali
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  if (usersSheet.getLastRow() < 2) {
    usersSheet.appendRow(['admin', '123456', 'SuperAdmin', '', '']);
  }
  Logger.log('Setup selesai. Sheet & akun default (admin / PIN 123456) sudah dibuat.');
}

/* ================= ROUTER ================= */
function doGet(e) {
  return jsonOutput({ ok: true, message: 'API Pondok aktif' });
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'login':        result = handleLogin(body); break;
      case 'verifyOtp':     result = handleVerifyOtp(body); break;
      case 'getData':       result = handleGetData(body); break;
      case 'getSantriDetail': result = handleGetSantriDetail(body); break;
      case 'saveSetoran':   result = handleSaveSetoran(body); break;
      case 'saveAbsen':     result = handleSaveAbsen(body); break;
      case 'uploadFile':    result = handleUploadFile(body); break;
      case 'sendWA':        result = handleSendWA(body); break;
      case 'saveKeuangan':  result = handleSaveKeuangan(body); break;
      case 'getKeuangan':   result = handleGetKeuangan(body); break;
      case 'getMonthlyProgress': result = handleGetMonthlyProgress(body); break;
      case 'getAbsenRekap': result = handleGetAbsenRekap(body); break;
      case 'getAbsen':      result = handleGetAbsen(body); break;          //FITUR BARU
      case 'savePresentasi': result = handleSavePresentasi(body); break;   //FITUR BARU
      case 'getPresentasiData': result = handleGetPresentasiData(body); break; //FITUR BARU
      default:              result = { ok: false, error: 'Action tidak dikenali: ' + body.action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return jsonOutput(result);
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================= HELPER SHEET ================= */
function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + name + ' (jalankan setupSheets() dulu)');
  return sheet;
}

function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i].every(c => c === '')) continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = values[i][idx]; });
    rows.push(obj);
  }
  return rows;
}

function appendRowFromObject(sheet, obj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => (obj[h] !== undefined ? obj[h] : ''));
  sheet.appendRow(row);
}

function findRowIndexById(sheet, idColName, idValue) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf(idColName);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(idValue)) return { rowIndex: i + 1, headers };
  }
  return null;
}

/* ================= LOGIN + OTP ================= */
function handleLogin(body) {
  const { username, pin } = body;
  if (!username || !pin) return { ok: false, error: 'Username dan PIN wajib diisi' };

  const users = sheetToObjects(getSheet(SHEET_USERS));
  const user = users.find(u => String(u.username).toLowerCase() === String(username).toLowerCase());
  if (!user) return { ok: false, error: 'Akun tidak ditemukan' };
  if (String(user.pin) !== String(pin)) return { ok: false, error: 'PIN salah' };

  const profile = {
    username: user.username,
    role: user.role,
    halaqoh: user.halaqoh || '',
    no_wa: user.no_wa || ''
  };

  if (ROLE_WAJIB_OTP.indexOf(user.role) !== -1) {
    if (!user.no_wa) return { ok: false, error: 'Nomor WA belum diset untuk akun ini, tidak bisa kirim OTP' };
    const otp = String(Math.floor(100000 + Math.random() * 900000));
    CacheService.getScriptCache().put('otp_' + user.username, otp, 300); // berlaku 5 menit
    const waResult = sendWaMessage(user.no_wa, `Kode OTP login Sistem Pondok kamu: ${otp} (berlaku 5 menit). Jangan bagikan kode ini ke siapa pun.`);
    return { ok: true, requiresOtp: true, message: 'OTP dikirim ke WA terdaftar', waSent: waResult.ok };
  }

  return { ok: true, requiresOtp: false, user: profile };
}

function handleVerifyOtp(body) {
  const { username, otp } = body;
  if (!username || !otp) return { ok: false, error: 'Username dan kode OTP wajib diisi' };
  const cached = CacheService.getScriptCache().get('otp_' + username);
  if (!cached) return { ok: false, error: 'Kode OTP kedaluwarsa, silakan login ulang' };
  if (String(cached) !== String(otp)) return { ok: false, error: 'Kode OTP salah' };
  CacheService.getScriptCache().remove('otp_' + username);

  const users = sheetToObjects(getSheet(SHEET_USERS));
  const user = users.find(u => String(u.username).toLowerCase() === String(username).toLowerCase());
  if (!user) return { ok: false, error: 'Akun tidak ditemukan' };

  return {
    ok: true,
    user: { username: user.username, role: user.role, halaqoh: user.halaqoh || '', no_wa: user.no_wa || '' }
  };
}

/* ================= DATA SANTRI ================= */
function handleGetData(body) {
  let rows = sheetToObjects(getSheet(SHEET_SANTRI));

  // Wali hanya boleh melihat data anaknya sendiri (dicocokkan lewat id santri yang dikirim client)
  if (body.role === 'Wali') {
    rows = body.id ? rows.filter(s => String(s.id) === String(body.id)) : [];
  }
  if (body.halaqoh) rows = rows.filter(s => String(s.halaqoh) === String(body.halaqoh));
  if (body.kelas) rows = rows.filter(s => String(s.kelas) === String(body.kelas));

  return { ok: true, data: rows };
}

function handleGetSantriDetail(body) {
  if (!body.id) return { ok: false, error: 'ID santri wajib diisi' };
  // Wali & Santri cuma boleh lihat datanya sendiri (username login = id santri)
  if ((body.role === 'Wali' || body.role === 'Santri') && String(body.requesterId) !== String(body.id)) {
    return { ok: false, error: 'Tidak punya akses ke data santri ini' };
  }
  const santri = sheetToObjects(getSheet(SHEET_SANTRI)).find(s => String(s.id) === String(body.id));
  if (!santri) return { ok: false, error: 'Santri tidak ditemukan' };
  const riwayat = sheetToObjects(getSheet(SHEET_SETORAN)).filter(s => String(s.id_santri) === String(body.id));
  return { ok: true, santri, riwayat };
}

/* ================= SETORAN (Qur'an / Kitab / Tahsin / Murajaah) ================= */
function handleSaveSetoran(body) {
  if (!body.id_santri || !body.jenis) return { ok: false, error: 'id_santri dan jenis wajib diisi' };
  appendRowFromObject(getSheet(SHEET_SETORAN), {
    id_santri: body.id_santri,
    tanggal: body.tanggal || todayStr(),
    jenis: body.jenis,
    jumlah: body.jumlah || '',
    nilai: body.nilai || '',
    catatan: body.catatan || '',
    bukti_url: body.bukti_url || '',
    ustadz: body.ustadz || ''
  });
  recalcSantriProgress(body.id_santri);
  return { ok: true, message: 'Setoran tersimpan' };
}

// Poin: Setor* +10, Murajaah +5, Telat -5, Bolos -20 (sesuai rumus yang diminta).
// "Setor Kitab" / "Setor Tahsin" diakumulasi ke kolom tercapai_kitab / tercapai_tahsin;
// "Setor" polos atau "Setor Quran" dianggap progres Qur'an (kolom tercapai utama).
function recalcSantriProgress(idSantri) {
  const santriSheet = getSheet(SHEET_SANTRI);
  const found = findRowIndexById(santriSheet, 'id', idSantri);
  if (!found) return;

  const riwayat = sheetToObjects(getSheet(SHEET_SETORAN)).filter(s => String(s.id_santri) === String(idSantri));
  let totalQuran = 0, totalKitab = 0, totalTahsin = 0, poin = 0;

  riwayat.forEach(s => {
    const jumlah = Number(s.jumlah) || 0;
    const jenis = String(s.jenis || '');
    if (jenis.indexOf('Setor') === 0) {
      poin += 10;
      if (jenis.indexOf('Kitab') !== -1) totalKitab += jumlah;
      else if (jenis.indexOf('Tahsin') !== -1) totalTahsin += jumlah;
      else totalQuran += jumlah;
    } else if (jenis === 'Murajaah') {
      poin += 5;
    } else if (jenis === 'Telat') {
      poin -= 5;
    } else if (jenis === 'Bolos') {
      poin -= 20;
    }
  });

  setCellByHeader(santriSheet, found, 'tercapai', totalQuran);
  setCellByHeader(santriSheet, found, 'tercapai_kitab', totalKitab);
  setCellByHeader(santriSheet, found, 'tercapai_tahsin', totalTahsin);
  setCellByHeader(santriSheet, found, 'poin', poin);
}

// Aman dipanggil walau kolom belum ada di sheet lama (skip diam-diam, tidak error)
function setCellByHeader(sheet, found, headerName, value) {
  const col = found.headers.indexOf(headerName);
  if (col === -1) return;
  sheet.getRange(found.rowIndex, col + 1).setValue(value);
}

/* ================= ABSEN — MODUL BARU (sheet Absensi terpisah) ================= */
//FITUR BARU — helper: ambil sheet Absensi, buat otomatis kalau belum ada
// (self-healing, tidak perlu jalankan ulang setupSheets di sheet yang sudah lama dipakai)
function getAbsensiSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_ABSENSI);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ABSENSI);
    const headers = ['tanggal', 'id_santri', 'status', 'catatan', 'input_by', 'input_role'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

//FITUR BARU — validasi akses: Ustadz cuma boleh untuk halaqoh-nya sendiri,
// Admin/SuperAdmin bebas semua halaqoh. Dipakai bareng modul Absen & Presentasi.
function checkHalaqohAccess(body, idSantri) {
  if (body.role === 'Admin' || body.role === 'SuperAdmin') return { ok: true };
  if (body.role !== 'Ustadz') return { ok: false, error: 'Role tidak punya akses ke modul ini' };

  const santri = sheetToObjects(getSheet(SHEET_SANTRI)).find(s => String(s.id) === String(idSantri));
  if (!santri) return { ok: false, error: 'Santri tidak ditemukan' };
  if (String(santri.halaqoh) !== String(body.halaqoh)) {
    return { ok: false, error: 'Ustadz hanya bisa mengakses santri di halaqoh-nya sendiri' };
  }
  return { ok: true };
}

//FITUR BARU — simpan/update absen 1 santri di 1 tanggal (idempotent per id_santri+tanggal)
function handleSaveAbsen(body) {
  if (!body.id_santri || !body.status || !body.tanggal) {
    return { ok: false, error: 'id_santri, status, dan tanggal wajib diisi' };
  }
  const akses = checkHalaqohAccess(body, body.id_santri);
  if (!akses.ok) return akses;

  const sheet = getAbsensiSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const tglCol = headers.indexOf('tanggal');
  const idCol = headers.indexOf('id_santri');
  const statusCol = headers.indexOf('status');
  const catatanCol = headers.indexOf('catatan');
  const byCol = headers.indexOf('input_by');
  const roleCol = headers.indexOf('input_role');

  for (let i = 1; i < values.length; i++) {
    if (formatDateOnly(values[i][tglCol]) === body.tanggal && String(values[i][idCol]) === String(body.id_santri)) {
      sheet.getRange(i + 1, statusCol + 1).setValue(body.status);
      sheet.getRange(i + 1, catatanCol + 1).setValue(body.catatan || '');
      sheet.getRange(i + 1, byCol + 1).setValue(body.username || '');
      sheet.getRange(i + 1, roleCol + 1).setValue(body.role || '');
      return { ok: true, message: 'Absen diperbarui' };
    }
  }

  sheet.appendRow([body.tanggal, body.id_santri, body.status, body.catatan || '', body.username || '', body.role || '']);
  return { ok: true, message: 'Absen tersimpan' };
}

//FITUR BARU — ambil data absen. Filter: tanggal (satu hari) ATAU mulai+selesai (rentang untuk rekap bulanan), + halaqoh.
function handleGetAbsen(body) {
  if (body.role === 'Ustadz' && !body.halaqoh) {
    return { ok: false, error: 'Halaqoh wajib diisi untuk role Ustadz' };
  }
  const sheet = getAbsensiSheet();
  let rows = sheetToObjects(sheet).map(r => ({ ...r, tanggal: formatDateOnly(r.tanggal) }));

  if (body.tanggal) rows = rows.filter(r => r.tanggal === body.tanggal);
  if (body.mulai) rows = rows.filter(r => r.tanggal >= body.mulai);
  if (body.selesai) rows = rows.filter(r => r.tanggal <= body.selesai);

  const santriMap = {};
  sheetToObjects(getSheet(SHEET_SANTRI)).forEach(s => { santriMap[String(s.id)] = s; });

  if (body.halaqoh) {
    rows = rows.filter(r => {
      const s = santriMap[String(r.id_santri)];
      return s && String(s.halaqoh) === String(body.halaqoh);
    });
  }

  rows = rows.map(r => {
    const s = santriMap[String(r.id_santri)] || {};
    return { ...r, nama: s.nama || '(tidak ditemukan)', kelas: s.kelas || '-', halaqoh: s.halaqoh || '-' };
  });

  return { ok: true, data: rows };
}

/* ================= ABSEN (LAMA — dipertahankan agar tidak memutus riwayat lama) ================= */
// Catatan: modul Absen yang aktif sekarang pakai sheet "Absensi" (lihat handleSaveAbsen/handleGetAbsen di atas).
// Fungsi di bawah ini legacy dari implementasi sebelumnya (nulis ke RiwayatSetoran jenis "Absen") — tidak lagi dipanggil dari absen.html versi baru.
function handleGetAbsenRekap(body) {
  const { mulai, selesai, halaqoh } = body;
  if (!mulai || !selesai) return { ok: false, error: 'mulai dan selesai wajib diisi' };

  const absenRows = sheetToObjects(getSheet(SHEET_SETORAN)).filter(r => r.jenis === 'Absen');
  const santriMap = {};
  sheetToObjects(getSheet(SHEET_SANTRI)).forEach(s => { santriMap[String(s.id)] = s; });

  const filtered = absenRows
    .map(r => ({ ...r, tanggal: formatDateOnly(r.tanggal) }))
    .filter(r => r.tanggal >= mulai && r.tanggal <= selesai)
    .filter(r => {
      if (!halaqoh) return true;
      const s = santriMap[String(r.id_santri)];
      return s && String(s.halaqoh) === String(halaqoh);
    })
    .map(r => {
      const s = santriMap[String(r.id_santri)] || {};
      return {
        id_santri: r.id_santri,
        nama: s.nama || '(tidak ditemukan)',
        kelas: s.kelas || '-',
        halaqoh: s.halaqoh || '-',
        tanggal: r.tanggal,
        status: r.nilai,
        catatan: r.catatan,
        ustadz: r.ustadz
      };
    });

  return { ok: true, data: filtered };
}

function formatDateOnly(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd');
  return String(value).slice(0, 10);
}

/* ================= UPLOAD FILE (Google Drive) ================= */
function handleUploadFile(body) {
  const { base64, filename, mimeType } = body;
  if (!base64 || !filename) return { ok: false, error: 'Data file tidak lengkap' };
  const folder = getOrCreateFolder(FOLDER_UPLOAD_NAME);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType || 'application/octet-stream', filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: file.getUrl(), id: file.getId() };
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

/* ================= KEUANGAN ================= */
function handleSaveKeuangan(body) {
  if (!body.jenis || !body.jumlah) return { ok: false, error: 'jenis dan jumlah wajib diisi' };
  appendRowFromObject(getSheet(SHEET_KEUANGAN), {
    tanggal: body.tanggal || todayStr(),
    jenis: body.jenis, // Pemasukan / Pengeluaran
    keterangan: body.keterangan || '',
    jumlah: body.jumlah,
    bukti: body.bukti || ''
  });
  return { ok: true, message: 'Transaksi tersimpan' };
}

function handleGetKeuangan(body) {
  return { ok: true, data: sheetToObjects(getSheet(SHEET_KEUANGAN)) };
}

/* ================= PRESENTASI PENCAPAIAN — MODUL BARU ================= */
//FITUR BARU — simpan hasil presentasi 1 santri (Setor Baru + Murajaah + Nilai) sebagai
// 1 baris di RiwayatSetoran dengan jenis "Presentasi". Idempotent per id_santri+tanggal
// (kalau sudah pernah disimpan di tanggal yang sama, baris lama diperbarui).
function handleSavePresentasi(body) {
  if (!body.id_santri || !body.tanggal) return { ok: false, error: 'id_santri dan tanggal wajib diisi' };
  const akses = checkHalaqohAccess(body, body.id_santri);
  if (!akses.ok) return akses;

  const catatanGabungan = `Murajaah: ${body.murajaah || 0} hlm.` + (body.catatan ? ' ' + body.catatan : '');

  const sheet = getSheet(SHEET_SETORAN);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id_santri');
  const tglCol = headers.indexOf('tanggal');
  const jenisCol = headers.indexOf('jenis');
  const jumlahCol = headers.indexOf('jumlah');
  const nilaiCol = headers.indexOf('nilai');
  const catatanCol = headers.indexOf('catatan');
  const ustadzCol = headers.indexOf('ustadz');

  for (let i = 1; i < values.length; i++) {
    if (
      String(values[i][idCol]) === String(body.id_santri) &&
      values[i][jenisCol] === 'Presentasi' &&
      formatDateOnly(values[i][tglCol]) === body.tanggal
    ) {
      sheet.getRange(i + 1, jumlahCol + 1).setValue(body.setor_baru || 0);
      sheet.getRange(i + 1, nilaiCol + 1).setValue(body.nilai || '');
      sheet.getRange(i + 1, catatanCol + 1).setValue(catatanGabungan);
      sheet.getRange(i + 1, ustadzCol + 1).setValue(body.username || '');
      return { ok: true, message: 'Presentasi diperbarui' };
    }
  }

  appendRowFromObject(sheet, {
    id_santri: body.id_santri,
    tanggal: body.tanggal,
    jenis: 'Presentasi',
    jumlah: body.setor_baru || 0,
    nilai: body.nilai || '',
    catatan: catatanGabungan,
    bukti_url: '',
    ustadz: body.username || ''
  });
  return { ok: true, message: 'Presentasi tersimpan' };
}

//FITUR BARU — data untuk grid presentasi: daftar santri (sudah difilter halaqoh + akses role)
// + data presentasi yang sudah diisi di tanggal terpilih (untuk prefill form).
function handleGetPresentasiData(body) {
  if (body.role === 'Ustadz' && !body.halaqoh) {
    return { ok: false, error: 'Halaqoh wajib diisi untuk role Ustadz' };
  }
  let santri = sheetToObjects(getSheet(SHEET_SANTRI));
  if (body.halaqoh) santri = santri.filter(s => String(s.halaqoh) === String(body.halaqoh));

  const tanggal = body.tanggal || todayStr();
  const presentasi = sheetToObjects(getSheet(SHEET_SETORAN))
    .filter(r => r.jenis === 'Presentasi' && formatDateOnly(r.tanggal) === tanggal);

  return { ok: true, santri, presentasi };
}

/* ================= WHATSAPP (Fonnte) ================= */
function handleSendWA(body) {
  if (!body.target || !body.message) return { ok: false, error: 'target dan message wajib diisi' };
  return sendWaMessage(body.target, body.message);
}

function sendWaMessage(target, message) {
  if (!FONNTE_TOKEN || FONNTE_TOKEN.indexOf('GANTI_DENGAN') === 0) {
    return { ok: false, error: 'FONNTE_TOKEN belum diisi di kode.gs' };
  }
  try {
    const response = UrlFetchApp.fetch('https://api.fonnte.com/send', {
      method: 'post',
      headers: { Authorization: FONNTE_TOKEN },
      payload: { target: target, message: message },
      muteHttpExceptions: true
    });
    return { ok: true, response: response.getContentText() };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/* ================= DASHBOARD: PROGRES BULANAN ================= */
// Menjumlahkan "jumlah" dari RiwayatSetoran (jenis Setor & Murajaah) per bulan,
// untuk 6 bulan terakhir. Dipakai grafik garis di dashboard.
function handleGetMonthlyProgress(body) {
  const riwayat = sheetToObjects(getSheet(SHEET_SETORAN))
    .filter(r => r.jenis === 'Setor' || r.jenis === 'Murajaah');

  const bulanMap = {}; // key: 'yyyy-MM' -> total
  const now = new Date();
  const order = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const key = Utilities.formatDate(d, Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM');
    bulanMap[key] = 0;
    order.push(key);
  }

  riwayat.forEach(r => {
    if (!r.tanggal) return;
    const d = new Date(r.tanggal);
    if (isNaN(d)) return;
    const key = Utilities.formatDate(d, Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM');
    if (bulanMap[key] !== undefined) bulanMap[key] += Number(r.jumlah) || 0;
  });

  const labels = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  const result = order.map(key => {
    const [y, m] = key.split('-');
    return { bulan: labels[Number(m) - 1] + ' ' + y, total: bulanMap[key] };
  });

  return { ok: true, data: result };
}

/* ================= UTIL ================= */
function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd');
}
