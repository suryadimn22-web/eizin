/**
 * e-Izin MAN 2 Kota Pekanbaru — Backend Google Apps Script
 * --------------------------------------------------------
 * Cara pakai singkat (lihat PANDUAN untuk detail):
 * 1. Buat Google Spreadsheet baru → Ekstensi → Apps Script.
 * 2. Hapus isi Code.gs bawaan, tempel seluruh file ini, simpan.
 * 3. Jalankan fungsi "setup" sekali (izinkan akses yang diminta).
 * 4. Isi sheet "Siswa" dan "Pengaturan" (terutama PIN Admin).
 * 5. Terapkan → Deployment baru → Aplikasi web
 *    (Jalankan sebagai: Saya, Akses: Siapa saja) → salin URL /exec.
 */

var SHEET_IZIN = 'Izin';
var SHEET_SISWA = 'Siswa';
var SHEET_SET = 'Pengaturan';
var SHEET_DETAIL = 'Detail_Siswa';
var SHEET_REKAP = 'Rekap';
var FOLDER_LAMPIRAN = 'e-Izin Lampiran';

var HEADERS_IZIN = [
  'ID', 'Kode Verifikasi', 'Dibuat', 'Sumber', 'Jenis Pemohon', 'Nama Pemohon', 'HP Pemohon',
  'Jenis Izin', 'Daftar Siswa', 'Jumlah Siswa', 'Tanggal Mulai', 'Tanggal Selesai',
  'Jam Keluar', 'Jam Kembali', 'Keperluan', 'Nama Kegiatan/Lomba', 'Tempat', 'Penyelenggara',
  'Pembina Pendamping', 'Lampiran', 'Status', 'Diputuskan Pada', 'Catatan Waka',
  'Kembali Aktual', 'Data (jangan diubah)'
];
var COL_DATA = HEADERS_IZIN.length; // kolom terakhir = JSON lengkap

var HEADERS_SISWA = ['NIS', 'Nama', 'Kelas'];
var HEADERS_DETAIL = ['ID', 'Tanggal Mulai', 'Bulan', 'NIS', 'Nama', 'Kelas', 'Jenis Izin', 'Jenis Pemohon', 'Status', 'Sumber'];

var DEFAULT_SET = [
  ['Nama Madrasah', 'MAN 2 Kota Pekanbaru'],
  ['Kop Baris 1', 'KEMENTERIAN AGAMA REPUBLIK INDONESIA'],
  ['Kop Baris 2', 'KANTOR KEMENTERIAN AGAMA KOTA PEKANBARU'],
  ['Kop Baris 3', 'MADRASAH ALIYAH NEGERI 2 KOTA PEKANBARU'],
  ['Alamat', '(isi alamat madrasah)'],
  ['Nama Penandatangan', 'Suryadi, S.E., M.M.'],
  ['NIP Penandatangan', '(isi NIP)'],
  ['Jabatan Penandatangan', 'Wakil Kepala Madrasah Bidang Kurikulum'],
  ['Kota Surat', 'Pekanbaru'],
  ['No WA Admin', ''],
  ['Email Notifikasi', ''],
  ['PIN Admin', '123456'],
  ['Kode Akses Pemohon', ''],
  ['Tanda Tangan', '']
];

// ---------------------------------------------------------------- setup

function setup() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEET_IZIN, HEADERS_IZIN);
  var siswa = ensureSheet_(ss, SHEET_SISWA, HEADERS_SISWA);
  ensureSheet_(ss, SHEET_DETAIL, HEADERS_DETAIL);
  ensureSheet_(ss, SHEET_REKAP, ['Rekap otomatis — dibuat ulang setiap ada perubahan data']);
  var set = ensureSheet_(ss, SHEET_SET, ['Pengaturan', 'Nilai']);
  if (set.getLastRow() < 2) {
    set.getRange(2, 1, DEFAULT_SET.length, 2).setValues(DEFAULT_SET);
    set.setColumnWidth(1, 220);
    set.setColumnWidth(2, 380);
  }
  if (siswa.getLastRow() < 2) {
    siswa.getRange(2, 1, 1, 3).setValues([['(contoh) 0001', '(contoh) Nama Siswa', 'XI-1']]);
  }
  // Kolom NIS & tanggal sebagai teks agar tidak berubah format
  siswa.getRange('A:A').setNumberFormat('@');
  rebuildRekap_();
  return 'Setup selesai';
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('e-Izin')
    .addItem('Siapkan / perbaiki sheet', 'setup')
    .addItem('Perbarui rekap', 'perbaruiRekap')
    .addToUi();
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e6f2ea');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ---------------------------------------------------------------- HTTP

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    switch (p.action) {
      case 'config': return json_({ ok: true, config: publicConfig_() });
      case 'siswa': return json_(apiSiswa_(p.kodeAkses));
      case 'cek': return json_(apiCek_(p.id, p.k));
      default: return json_({ ok: true, app: 'e-Izin MAN 2', versi: 1 });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'Format data tidak valid' });
  }
  try {
    switch (body.action) {
      case 'ajukan': return json_(apiAjukan_(body));
      case 'login': return json_(apiLogin_(body));
      case 'list': return json_(apiList_(body));
      case 'sync': return json_(apiSync_(body));
      default: return json_({ ok: false, error: 'Aksi tidak dikenal' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------- settings

function settings_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SET);
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  vals.forEach(function (r) { if (r[0]) out[String(r[0]).trim()] = r[1] === null ? '' : String(r[1]); });
  return out;
}

function setSetting_(key, value) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SET);
  var last = sh.getLastRow();
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === key) { sh.getRange(i + 2, 2).setValue(value); return; }
    }
  }
  sh.appendRow([key, value]);
}

function publicConfig_() {
  var s = settings_();
  return {
    namaMadrasah: s['Nama Madrasah'] || '',
    kop1: s['Kop Baris 1'] || '', kop2: s['Kop Baris 2'] || '', kop3: s['Kop Baris 3'] || '',
    alamat: s['Alamat'] || '',
    penandatangan: s['Nama Penandatangan'] || '',
    nip: s['NIP Penandatangan'] || '',
    jabatan: s['Jabatan Penandatangan'] || '',
    kota: s['Kota Surat'] || '',
    waAdmin: s['No WA Admin'] || '',
    perluKode: !!String(s['Kode Akses Pemohon'] || '').trim(),
    ttd: s['Tanda Tangan'] || ''
  };
}

function checkPin_(pin) {
  var real = String(settings_()['PIN Admin'] || '').trim();
  if (!real) throw new Error('PIN Admin belum diisi di sheet Pengaturan');
  if (String(pin || '').trim() !== real) throw new Error('PIN salah');
}

function checkKodeAkses_(kode) {
  var real = String(settings_()['Kode Akses Pemohon'] || '').trim();
  if (real && String(kode || '').trim().toLowerCase() !== real.toLowerCase()) {
    throw new Error('Kode akses pemohon salah');
  }
}

// ---------------------------------------------------------------- data access

function readSiswa_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SISWA);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(function (r) { return String(r[1]).trim() && String(r[1]).indexOf('(contoh)') !== 0; })
    .map(function (r) { return { nis: String(r[0]).trim(), nama: String(r[1]).trim(), kelas: String(r[2]).trim() }; });
}

function readIzin_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_IZIN);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, COL_DATA).getValues();
  for (var i = 0; i < vals.length; i++) {
    var raw = vals[i][COL_DATA - 1];
    if (!raw) continue;
    try {
      var rec = JSON.parse(raw);
      rec._row = i + 2;
      out.push(rec);
    } catch (err) { /* baris rusak dilewati */ }
  }
  return out;
}

function toRow_(r) {
  var siswa = (r.siswa || []).map(function (s) { return s.nama + ' (' + (s.kelas || '-') + ')'; }).join('; ');
  var clean = {};
  Object.keys(r).forEach(function (k) { if (k.charAt(0) !== '_' && k !== 'lampiranData') clean[k] = r[k]; });
  return [
    r.id, r.kode, r.dibuat || '', r.sumber || '', r.jenisPemohon || '', r.namaPemohon || '', r.hpPemohon || '',
    r.jenisIzin || '', siswa, (r.siswa || []).length, r.tglMulai || '', r.tglSelesai || '',
    r.jamKeluar || '', r.jamKembali || '', r.keperluan || '', r.namaKegiatan || '', r.tempat || '',
    r.penyelenggara || '', r.pembina || '', r.lampiran || '', r.status || '', r.diputuskan || '',
    r.catatan || '', r.kembaliAktual || '', JSON.stringify(clean)
  ];
}

function writeIzin_(rec) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_IZIN);
  var row = toRow_(rec);
  if (rec._row) {
    sh.getRange(rec._row, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
    rec._row = sh.getLastRow();
  }
  // simpan kolom teks apa adanya (tanggal/jam tidak dikonversi Sheets)
  sh.getRange(rec._row, 11, 1, 4).setNumberFormat('@').setValues([[rec.tglMulai || '', rec.tglSelesai || '', rec.jamKeluar || '', rec.jamKembali || '']]);
}

function strip_(rec) {
  var o = {};
  Object.keys(rec).forEach(function (k) { if (k.charAt(0) !== '_' && k !== 'lampiranData') o[k] = rec[k]; });
  return o;
}

// ---------------------------------------------------------------- helpers

var ALPHA_ = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
function rand_(n) {
  var s = '';
  for (var i = 0; i < n; i++) s += ALPHA_.charAt(Math.floor(Math.random() * ALPHA_.length));
  return s;
}
function newId_() {
  var d = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd');
  return 'IZN-' + d + '-' + rand_(4);
}
function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Jakarta', "yyyy-MM-dd'T'HH:mm:ss");
}
function str_(v, max) {
  v = v === undefined || v === null ? '' : String(v).trim();
  return v.length > max ? v.substring(0, max) : v;
}

var JENIS_PEMOHON = ['Wali Kelas', 'Pembina', 'Koordinator Kelas Program', 'Langsung (Waka)', 'Siswa'];
var JENIS_IZIN = ['Keluar Sekolah Sementara', 'Pulang Lebih Awal', 'Mengikuti Lomba/Kegiatan', 'Lainnya'];
var STATUS = ['Menunggu', 'Disetujui', 'Ditolak'];

function sanitize_(r, fromAdmin) {
  var siswa = Array.isArray(r.siswa) ? r.siswa : [];
  siswa = siswa.slice(0, 60).map(function (s) {
    return { nis: str_(s.nis, 30), nama: str_(s.nama, 100), kelas: str_(s.kelas, 30) };
  }).filter(function (s) { return s.nama; });
  var out = {
    id: /^IZN-\d{6}-[A-Z0-9]{4}$/.test(String(r.id || '')) ? String(r.id) : newId_(),
    kode: /^[A-Z0-9]{6}$/.test(String(r.kode || '')) ? String(r.kode) : rand_(6),
    dibuat: str_(r.dibuat, 25) || now_(),
    sumber: fromAdmin ? (str_(r.sumber, 30) || 'Langsung') : 'Online',
    jenisPemohon: JENIS_PEMOHON.indexOf(r.jenisPemohon) >= 0 ? r.jenisPemohon : 'Wali Kelas',
    namaPemohon: str_(r.namaPemohon, 100),
    hpPemohon: str_(r.hpPemohon, 20),
    jenisIzin: JENIS_IZIN.indexOf(r.jenisIzin) >= 0 ? r.jenisIzin : 'Lainnya',
    siswa: siswa,
    tglMulai: str_(r.tglMulai, 10),
    tglSelesai: str_(r.tglSelesai, 10),
    jamKeluar: str_(r.jamKeluar, 5),
    jamKembali: str_(r.jamKembali, 5),
    keperluan: str_(r.keperluan, 1000),
    namaKegiatan: str_(r.namaKegiatan, 200),
    tempat: str_(r.tempat, 200),
    penyelenggara: str_(r.penyelenggara, 200),
    pembina: str_(r.pembina, 100),
    lampiran: str_(r.lampiran, 300),
    status: fromAdmin && STATUS.indexOf(r.status) >= 0 ? r.status : 'Menunggu',
    diputuskan: fromAdmin ? str_(r.diputuskan, 25) : '',
    catatan: fromAdmin ? str_(r.catatan, 500) : '',
    kembaliAktual: fromAdmin ? str_(r.kembaliAktual, 25) : ''
  };
  if (!out.namaPemohon) throw new Error('Nama pemohon wajib diisi');
  if (!out.siswa.length) throw new Error('Minimal satu siswa harus dipilih');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(out.tglMulai)) throw new Error('Tanggal mulai tidak valid');
  if (!fromAdmin && (!/^\d{2}:\d{2}$/.test(out.jamKeluar) || !/^\d{2}:\d{2}$/.test(out.jamKembali))) {
    throw new Error('Jam keluar dan jam kembali wajib diisi');
  }
  return out;
}

function saveLampiran_(dataUrl, id) {
  var m = /^data:(image\/(jpeg|png|webp));base64,(.+)$/.exec(String(dataUrl || ''));
  if (!m) return '';
  if (m[3].length > 3000000) throw new Error('Lampiran terlalu besar');
  var it = DriveApp.getFoldersByName(FOLDER_LAMPIRAN);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_LAMPIRAN);
  var blob = Utilities.newBlob(Utilities.base64Decode(m[3]), m[1], id + '.' + (m[2] === 'jpeg' ? 'jpg' : m[2]));
  var file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) { /* sebagian domain sekolah melarang berbagi publik — file tetap tersimpan */ }
  return file.getUrl();
}

function notify_(rec) {
  var email = String(settings_()['Email Notifikasi'] || '').trim();
  if (!email) return;
  var names = rec.siswa.map(function (s) { return '- ' + s.nama + ' (' + s.kelas + ')'; }).join('\n');
  MailApp.sendEmail(email, '[e-Izin] Pengajuan baru: ' + rec.jenisIzin + ' — ' + rec.namaPemohon,
    'Pengajuan izin baru menunggu persetujuan.\n\n' +
    'ID: ' + rec.id + '\nPemohon: ' + rec.namaPemohon + ' (' + rec.jenisPemohon + ')\n' +
    'Jenis izin: ' + rec.jenisIzin + '\nTanggal: ' + rec.tglMulai + (rec.tglSelesai ? ' s.d. ' + rec.tglSelesai : '') + '\n' +
    'Keperluan: ' + rec.keperluan + '\n\nSiswa:\n' + names + '\n\nBuka aplikasi e-Izin untuk menyetujui/menolak.');
}

// ---------------------------------------------------------------- API handlers

function apiSiswa_(kodeAkses) {
  checkKodeAkses_(kodeAkses);
  return { ok: true, siswa: readSiswa_() };
}

function apiCek_(id, k) {
  if (!id || !k) return { ok: false, error: 'ID dan kode wajib diisi' };
  var recs = readIzin_();
  for (var i = 0; i < recs.length; i++) {
    if (recs[i].id === id) {
      if (recs[i].kode !== String(k).trim().toUpperCase()) return { ok: false, error: 'Kode verifikasi tidak cocok' };
      return { ok: true, izin: strip_(recs[i]), config: publicConfig_() };
    }
  }
  return { ok: false, error: 'Izin tidak ditemukan' };
}

function apiAjukan_(body) {
  checkKodeAkses_(body.kodeAkses);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var input = body.izin || {};
    var recs = readIzin_();
    // idempoten: pengajuan dari antrean offline bisa terkirim dua kali
    for (var i = 0; i < recs.length; i++) {
      if (input.id && recs[i].id === input.id) return { ok: true, id: recs[i].id, kode: recs[i].kode, duplikat: true };
    }
    var rec = sanitize_(input, false);
    if (input.lampiranData) {
      try { rec.lampiran = saveLampiran_(input.lampiranData, rec.id); }
      catch (err) { rec.lampiran = ''; rec.keperluan = (rec.keperluan + ' [foto lampiran gagal disimpan]').substring(0, 1000); }
    }
    writeIzin_(rec);
    rebuildRekap_();
    try { notify_(rec); } catch (err) { /* notifikasi opsional */ }
    return { ok: true, id: rec.id, kode: rec.kode };
  } finally {
    lock.releaseLock();
  }
}

function apiLogin_(body) {
  checkPin_(body.pin);
  return { ok: true, config: publicConfig_() };
}

function apiList_(body) {
  checkPin_(body.pin);
  return { ok: true, izin: readIzin_().map(strip_), siswa: readSiswa_(), config: publicConfig_() };
}

/**
 * ops: [{type:'upsert', izin:{...}} | {type:'putus', id, status, catatan, waktu}
 *       | {type:'kembali', id, waktu} | {type:'ttd', dataUrl}]
 */
function apiSync_(body) {
  checkPin_(body.pin);
  var ops = Array.isArray(body.ops) ? body.ops : [];
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  var hasil = [];
  try {
    var recs = readIzin_();
    var byId = {};
    recs.forEach(function (r) { byId[r.id] = r; });
    var changed = false;
    ops.forEach(function (op, idx) {
      try {
        if (op.type === 'upsert') {
          var rec = sanitize_(op.izin || {}, true);
          if (byId[rec.id]) rec._row = byId[rec.id]._row;
          writeIzin_(rec);
          byId[rec.id] = rec;
          changed = true;
        } else if (op.type === 'putus') {
          var r = byId[op.id];
          if (!r) throw new Error('ID tidak ditemukan: ' + op.id);
          if (STATUS.indexOf(op.status) < 0) throw new Error('Status tidak valid');
          r.status = op.status;
          r.catatan = str_(op.catatan, 500);
          r.diputuskan = str_(op.waktu, 25) || now_();
          writeIzin_(r);
          changed = true;
        } else if (op.type === 'kembali') {
          var k = byId[op.id];
          if (!k) throw new Error('ID tidak ditemukan: ' + op.id);
          k.kembaliAktual = str_(op.waktu, 25) || now_();
          writeIzin_(k);
          changed = true;
        } else if (op.type === 'ttd') {
          var d = String(op.dataUrl || '');
          if (d && !/^data:image\/png;base64,/.test(d)) throw new Error('Format tanda tangan tidak valid');
          if (d.length > 49000) throw new Error('Gambar tanda tangan terlalu besar');
          setSetting_('Tanda Tangan', d);
        }
        hasil.push({ i: idx, ok: true });
      } catch (err) {
        hasil.push({ i: idx, ok: false, error: String(err && err.message || err) });
      }
    });
    if (changed) rebuildRekap_();
  } finally {
    lock.releaseLock();
  }
  return { ok: true, hasil: hasil, izin: readIzin_().map(strip_), siswa: readSiswa_(), config: publicConfig_() };
}

// ---------------------------------------------------------------- rekap

var BULAN_ = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

/** Bisa dijalankan manual dari menu e-Izin. */
function perbaruiRekap() { rebuildRekap_(); }

function rebuildRekap_() {
  var ss = SpreadsheetApp.getActive();
  var recs = readIzin_();

  // Detail per siswa (1 baris = 1 siswa dalam 1 izin) — mudah untuk filter/pivot
  var det = ss.getSheetByName(SHEET_DETAIL) || ensureSheet_(ss, SHEET_DETAIL, HEADERS_DETAIL);
  var rows = [];
  recs.forEach(function (r) {
    var bln = r.tglMulai ? r.tglMulai.substring(0, 7) : '';
    (r.siswa || []).forEach(function (s) {
      rows.push([r.id, r.tglMulai, bln, s.nis, s.nama, s.kelas, r.jenisIzin, r.jenisPemohon, r.status, r.sumber]);
    });
  });
  if (det.getLastRow() > 1) det.getRange(2, 1, det.getLastRow() - 1, HEADERS_DETAIL.length).clearContent();
  if (rows.length) {
    det.getRange(2, 1, rows.length, HEADERS_DETAIL.length).setNumberFormat('@').setValues(rows);
  }

  // Ringkasan
  var sh = ss.getSheetByName(SHEET_REKAP) || ss.insertSheet(SHEET_REKAP);
  sh.clear();
  var out = [];
  out.push(['REKAP IZIN SISWA', '', '', '', '']);
  out.push(['Diperbarui: ' + now_().replace('T', ' '), '', '', '', '']);
  out.push(['', '', '', '', '']);
  var st = { Menunggu: 0, Disetujui: 0, Ditolak: 0 };
  recs.forEach(function (r) { st[r.status] = (st[r.status] || 0) + 1; });
  out.push(['Status pengajuan', 'Jumlah', '', '', '']);
  STATUS.forEach(function (s) { out.push([s, st[s] || 0, '', '', '']); });
  out.push(['', '', '', '', '']);

  // Per bulan x jenis izin (yang disetujui, dihitung per siswa)
  out.push(['Izin disetujui per bulan (jumlah siswa)'].concat(JENIS_IZIN));
  var perBulan = {};
  rows.forEach(function (d) {
    if (d[8] !== 'Disetujui') return;
    var key = d[2];
    perBulan[key] = perBulan[key] || {};
    perBulan[key][d[6]] = (perBulan[key][d[6]] || 0) + 1;
  });
  Object.keys(perBulan).sort().forEach(function (k) {
    var parts = k.split('-');
    var label = BULAN_[Number(parts[1]) - 1] + ' ' + parts[0];
    out.push([label].concat(JENIS_IZIN.map(function (j) { return perBulan[k][j] || 0; })));
  });
  out.push(['', '', '', '', '']);

  // Siswa paling sering izin
  out.push(['Siswa paling sering izin (disetujui)', 'Kelas', 'Total', 'Keluar/Pulang', 'Lomba/Kegiatan']);
  var perSiswa = {};
  rows.forEach(function (d) {
    if (d[8] !== 'Disetujui') return;
    var key = (d[3] || '') + '|' + d[4];
    var o = perSiswa[key] = perSiswa[key] || { nama: d[4], kelas: d[5], total: 0, keluar: 0, lomba: 0 };
    o.total++;
    if (d[6] === 'Mengikuti Lomba/Kegiatan') o.lomba++; else o.keluar++;
  });
  Object.keys(perSiswa).map(function (k) { return perSiswa[k]; })
    .sort(function (a, b) { return b.total - a.total; }).slice(0, 30)
    .forEach(function (o) { out.push([o.nama, o.kelas, o.total, o.keluar, o.lomba]); });

  var width = 5;
  out = out.map(function (r) { while (r.length < width) r.push(''); return r.slice(0, width); });
  sh.getRange(1, 1, out.length, width).setValues(out);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sh.setColumnWidth(1, 300);
}
