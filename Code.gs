/**
 * ============================================================
 * SISTEM MANAJEMEN SARAN PERBAIKAN - BACKEND (Google Apps Script)
 * ============================================================
 * Deploy sebagai Web App (Execute as: Me, Access: Anyone with link
 * atau Anyone within organization jika pakai Google Workspace).
 *
 * Struktur data (Google Spreadsheet - 1 file, 2 sheet utama):
 *
 * SHEET "Kunjungan" (1 baris = 1 unit yang masuk bengkel)
 * A: ID_Kunjungan
 * B: Timestamp
 * C: Tanggal_Masuk
 * D: No_Polisi
 * E: Nama_Customer
 * F: No_HP_Customer
 * G: SA (Service Advisor)
 * H: Teknisi
 * I: Status  (Menunggu Foreman / Siap Report / Report Terkirim / Follow Up / Deal / Tidak Deal)
 * J: PDF_URL
 * K: Tanggal_FollowUp_Rencana  (otomatis = Tanggal_Masuk + 3 hari, digeser +14 hari tiap Tidak Deal)
 * L: Status_FollowUp  (Belum / Tidak Deal / Deal - Belum Datang / Deal - Sudah Datang)
 * M: Catatan_FollowUp
 * N: Tanggal_FollowUp_Aktual
 * ... (kolom-kolom lain: lihat HEADERS.Kunjungan di bawah - daftar ini tidak diperbarui
 *      untuk tiap kolom baru, HEADERS adalah sumber kebenaran)
 * Perlu_FollowUp  (Ya/Tidak - dihitung saat Foreman kirim report: ada saran Jasa/Parts atau tidak)
 * Alasan_Tidak_Deal, Tanggal_Rencana_Datang, Jam_Rencana_Datang,
 * Tanggal_Kedatangan_Aktual, Nomor_PKB  (alur follow up SA, lihat updateRow dari frontend)
 *
 * SHEET "Item_Saran" (1 baris = 1 komponen yang disarankan)
 * A: ID_Item
 * B: ID_Kunjungan   (FK ke sheet Kunjungan)
 * C: Nama_Komponen        (wajib - Teknisi)
 * D: Qty                  (wajib - Teknisi)
 * E: Keterangan_Teknisi   (opsional - Teknisi)
 * F: Nomor_Part           (diisi Foreman)
 * G: Estimasi_Harga       (diisi Foreman)
 * H: Ketersediaan_Part    (diisi Foreman)
 * I: Keterangan_Partman   (opsional - diisi Foreman)
 * J: Foto_URL             (opsional, link Drive)
 * K: Diisi_Teknisi_At
 * L: Diisi_Partman_At
 * ============================================================
 */

const SPREADSHEET_ID = '1_JYeu0uYI1CxLA2Y5EMFDFNFaCZnhibG_--o-YGRRqA'; // ID Google Sheet (database)
const DRIVE_FOLDER_ID = '1A6VLdeox-bhZGS-u9XsyfOnOfTF41viz'; // Folder khusus foto komponen
const PDF_TEMPLATE_ID = '1-NkoCuTPNBP0iYoJBXoLW2BCAcNvK9LEg61PJ_ZqYx0'; // Template dokumen report Foreman->SA
const CODE_VERSION = 'v33-dashboard-perf-fix'; // Ganti tiap perubahan, dipakai action=version untuk cek deployment

// Header wajib per sheet - dipakai untuk memvalidasi/memulihkan row 1 setiap sheet diakses,
// supaya baris data tidak pernah tersalah-baca sebagai header (lihat ensureHeader()).
const HEADERS = {
  Kunjungan: ['ID_Kunjungan','Timestamp','Tanggal_Masuk','No_Polisi','Nama_Customer',
    'No_HP_Customer','SA','Teknisi','Status','PDF_URL','Tanggal_FollowUp_Rencana',
    'Status_FollowUp','Catatan_FollowUp','Tanggal_FollowUp_Aktual','Tipe_Mobil','Tipe_Service',
    'Pekerjaan','Request_Customer',
    'Aki_Status','Aki_Keterangan','Aki_Harga',
    'Ban_Status','Ban_Keterangan','Ban_Merk1','Ban_Harga1','Ban_Merk2','Ban_Harga2',
    'KampasRem_Status','KampasRem_Keterangan','KampasRem_Harga',
    'Wiper_Status','Wiper_Keterangan','Wiper_Harga',
    'Lampu_Status','Lampu_Keterangan','Lampu_Harga',
    'Mobil_Hybrid','SBE_50K_100K','Mobil_2_5_Tahun','HHC','HHC_Hasil',
    'SSC_Terlibat',
    'UjiEmisi_Status','UjiEmisi_Foto_URL',
    'Nama_Foreman','FollowUp_Dikonfirmasi',
    'Perlu_FollowUp','Alasan_Tidak_Deal','Tanggal_Rencana_Datang','Jam_Rencana_Datang',
    'Tanggal_Kedatangan_Aktual','Nomor_PKB'],
  // "Nama Parts" di Saran Perbaikan - Harga_Satuan_Teknisi adalah estimasi harga part dari
  // Teknisi sendiri saat input awal, terpisah dari Estimasi_Harga yang diisi Partman belakangan.
  Item_Saran: ['ID_Item','ID_Kunjungan','Nama_Komponen','Qty','Keterangan_Teknisi',
    'Nomor_Part','Estimasi_Harga','Ketersediaan_Part','Keterangan_Partman','Foto_URL',
    'Diisi_Teknisi_At','Diisi_Partman_At','Harga_Satuan_Teknisi'],
  // "Nama Jasa" di Saran Perbaikan - estimasi biaya jasa/pemasangan, terpisah dari harga part.
  Item_Jasa: ['ID_Jasa','ID_Kunjungan','Nama_Jasa','Waktu','Harga_Satuan','Keterangan','Diisi_Teknisi_At'],
  // Daftar SSC per kunjungan (1 kunjungan bisa ikut beberapa campaign sekaligus).
  Item_SSC: ['ID_SSC','ID_Kunjungan','SSC','Status','Alasan','Diisi_Teknisi_At'],
  // Daftar Technical Information per kunjungan (1 kunjungan bisa punya beberapa baris).
  Item_TechnicalInfo: ['ID_TI','ID_Kunjungan','Technical_Information','Status','Alasan','Diisi_Teknisi_At'],
  // Setiap kolom = 1 daftar pilihan dropdown (baris di bawah header = isi pilihan).
  // Kolom bisa punya jumlah baris berbeda-beda, sel kosong akan diabaikan.
  MasterData: ['Tipe_Mobil','Tipe_Service','Service_Advisor','Teknisi','SSC','Alasan_SSC',
    'Technical_Information','Alasan_Technical','Foreman']
};

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function getSheet(name) {
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  ensureHeader(sh, HEADERS[name]);
  return sh;
}

/** Pastikan row 1 selalu berisi header yang benar. Jika hilang/salah (mis. sheet
 *  baru langsung diisi data tanpa setupSheets()), sisipkan header baru di row 1
 *  dan dorong data yang sudah ada ke bawah, alih-alih membiarkan baris data
 *  tersalah-baca sebagai header oleh rowToObj(). Kalau header lama masih valid
 *  dan cuma kurang kolom baru di akhir (skema bertambah), header cukup
 *  diperluas di tempat tanpa menggeser baris data manapun. */
function ensureHeader(sh, headers) {
  if (!headers) return;
  const lastRow = sh.getLastRow();
  const firstRow = lastRow > 0 ? sh.getRange(1, 1, 1, headers.length).getValues()[0] : [];
  const isCorrect = headers.every((h, i) => firstRow[i] === h);
  if (isCorrect) return;

  const isSafeExtension = firstRow.length > 0 &&
    headers.every((h, i) => firstRow[i] === h || firstRow[i] === '' || firstRow[i] === undefined);
  if (isSafeExtension) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  if (lastRow > 0) sh.insertRowBefore(1);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
}

/** Setup awal - jalankan sekali manual dari editor Apps Script.
 *  Sebenarnya opsional sekarang karena getSheet() otomatis memvalidasi/memulihkan
 *  header setiap kali dipanggil, tapi tetap disediakan agar sesuai langkah di README. */
function setupSheets() {
  getSheet('Kunjungan');
  getSheet('Item_Saran');
  getSheet('Item_Jasa');
  getSheet('Item_SSC');
  getSheet('Item_TechnicalInfo');
  getSheet('MasterData');
}

/** ============ ENTRY POINTS ============ */
function doGet(e) {
  const action = e.parameter.action;
  try {
    let result;
    switch (action) {
      case 'listKunjungan': result = listKunjungan(e.parameter); break;
      case 'getKunjungan': result = getKunjunganDetail(e.parameter.id); break;
      case 'dashboard': result = getDashboard(e.parameter); break;
      case 'masterData': result = getMasterData(); break;
      case 'version': result = { version: CODE_VERSION }; break;
      default: result = { error: 'Unknown action' };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  try {
    let result;
    switch (action) {
      case 'createKunjungan': result = createKunjungan(body); break;
      case 'updateKunjungan': result = updateKunjungan(body); break;
      case 'addItemSaran': result = addItemSaran(body); break;
      case 'addItemJasa': result = addItemJasa(body); break;
      case 'addItemSSC': result = addItemSSC(body); break;
      case 'addItemTechnicalInfo': result = addItemTechnicalInfo(body); break;
      case 'updatePartman': result = updatePartmanItem(body); break;
      case 'updateRow': result = updateRowByField(body.sheet, body.id, body.fields); break;
      case 'deleteRow': result = deleteRowByField(body.sheet, body.id); break;
      case 'uploadFoto': result = uploadFoto(body); break;
      case 'generateReport': result = generateReport(body.idKunjungan, body.namaForeman, body.printedBy); break;
      case 'updateFollowUp': result = updateFollowUp(body); break;
      default: result = { error: 'Unknown action' };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** ============ TEKNISI: buat kunjungan baru + item awal ============ */
function createKunjungan(body) {
  const sh = getSheet('Kunjungan');
  const id = 'KJ-' + new Date().getTime();
  const now = new Date();
  const followUpDate = new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000);
  const p = body.proses || {};
  const aki = p.aki || {}, ban = p.ban || {}, kampasRem = p.kampasRem || {}, wiper = p.wiper || {}, lampu = p.lampu || {};
  const hhc = body.hybridHealthCheck || {};
  const ujiEmisi = body.ujiEmisi || {};
  sh.appendRow([id, now, body.tanggalMasuk || now, body.noPolisi, body.namaCustomer,
    body.noHp || '', body.sa, body.teknisi, 'Menunggu Foreman', '', followUpDate, 'Belum', '', '',
    body.tipeMobil || '', body.tipeService || '',
    p.pekerjaan || '', p.requestCustomer || '',
    aki.status || 'OK', aki.keterangan || '', aki.harga || '',
    ban.status || 'OK', ban.keterangan || '', ban.merk1 || '', ban.harga1 || '', ban.merk2 || '', ban.harga2 || '',
    kampasRem.status || 'OK', kampasRem.keterangan || '', kampasRem.harga || '',
    wiper.status || 'OK', wiper.keterangan || '', wiper.harga || '',
    lampu.status || 'OK', lampu.keterangan || '', lampu.harga || '',
    hhc.mobilHybrid || 'No', hhc.sbe50k100k || '', hhc.mobil2to5Tahun || '', hhc.hhc || '', hhc.hasil || '',
    body.sscTerlibat || 'Tidak Terlibat',
    ujiEmisi.status || 'Tidak Lulus/Belum Uji Emisi/Kadaluarsa', ujiEmisi.fotoUrl || '',
    '', 'Tidak']); // Nama_Foreman (diisi Foreman nanti), FollowUp_Dikonfirmasi (diisi SA nanti)

  // item-item Parts (wajib: namaKomponen, qty; opsional: keterangan, hargaSatuan, fotoUrl)
  if (body.items && body.items.length) {
    body.items.forEach(it => addItemSaran({ idKunjungan: id, ...it }));
  }
  // item-item Jasa (wajib: namaJasa; opsional: waktu, hargaSatuan, keterangan)
  if (body.itemsJasa && body.itemsJasa.length) {
    body.itemsJasa.forEach(it => addItemJasa({ idKunjungan: id, ...it }));
  }
  // item-item SSC (wajib: ssc; opsional: status, alasan)
  if (body.itemsSSC && body.itemsSSC.length) {
    body.itemsSSC.forEach(it => addItemSSC({ idKunjungan: id, ...it }));
  }
  // item-item Technical Information (wajib: technicalInformation; opsional: status, alasan)
  if (body.itemsTechnicalInfo && body.itemsTechnicalInfo.length) {
    body.itemsTechnicalInfo.forEach(it => addItemTechnicalInfo({ idKunjungan: id, ...it }));
  }
  return { success: true, idKunjungan: id };
}

/** Update sebagian field kunjungan yang sudah ada, dicari lewat nama kolom di header
 *  (bukan nomor kolom tetap) - supaya aman dipakai untuk field mana pun tanpa perlu
 *  hardcode index tiap kali skema Kunjungan bertambah. Dipakai oleh Foreman untuk
 *  mengedit Data Unit/Customer, Proses Service, dan section lain nantinya. */
function updateKunjungan(body) {
  if (!body.idKunjungan || !body.fields) {
    return { error: 'idKunjungan dan fields wajib diisi' };
  }
  const sh = getSheet('Kunjungan');
  const data = sh.getDataRange().getValues();
  const header = data[0];
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === body.idKunjungan) {
      Object.keys(body.fields).forEach(key => {
        const col = header.indexOf(key);
        if (col !== -1) sh.getRange(r + 1, col + 1).setValue(body.fields[key]);
      });
      return { success: true };
    }
  }
  return { error: 'Kunjungan tidak ditemukan' };
}

/** Update sebagian field 1 baris di sheet mana pun (dicari lewat ID di kolom A),
 *  field dicari lewat nama header supaya tidak perlu hardcode nomor kolom.
 *  Dipakai Foreman untuk full-edit Item_Saran/Item_Jasa/dst tanpa perlu fungsi
 *  update terpisah per sheet. sheetName divalidasi terhadap HEADERS supaya
 *  tidak bisa dipakai untuk menyentuh sheet sembarangan. */
// Kolom yang HARUS disimpan sebagai teks polos, bukan dibiarkan Sheets auto-parse
// (mis. "14:00" berubah jadi serial waktu 1899, atau "01234" kehilangan leading zero).
const TEXT_ONLY_FIELDS = ['Jam_Rencana_Datang', 'Nomor_PKB'];
function updateRowByField(sheetName, idValue, fields) {
  if (!HEADERS[sheetName]) return { error: 'Sheet tidak dikenal: ' + sheetName };
  if (!idValue || !fields) return { error: 'id dan fields wajib diisi' };
  const sh = getSheet(sheetName);
  const data = sh.getDataRange().getValues();
  const header = data[0];
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === idValue) {
      Object.keys(fields).forEach(key => {
        const col = header.indexOf(key);
        if (col === -1) return;
        const range = sh.getRange(r + 1, col + 1);
        if (TEXT_ONLY_FIELDS.indexOf(key) !== -1) range.setNumberFormat('@');
        range.setValue(fields[key]);
      });
      return { success: true };
    }
  }
  return { error: 'Baris tidak ditemukan' };
}

/** Hapus 1 baris dari sheet mana pun (dicari lewat ID di kolom A). */
function deleteRowByField(sheetName, idValue) {
  if (!HEADERS[sheetName]) return { error: 'Sheet tidak dikenal: ' + sheetName };
  if (!idValue) return { error: 'id wajib diisi' };
  const sh = getSheet(sheetName);
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === idValue) {
      sh.deleteRow(r + 1);
      return { success: true };
    }
  }
  return { error: 'Baris tidak ditemukan' };
}

/** Teknisi tambah 1 item part (Nama Parts) ke kunjungan yang sudah ada */
function addItemSaran(body) {
  if (!body.namaKomponen || !body.qty) {
    return { error: 'Nama komponen dan Qty wajib diisi' };
  }
  const sh = getSheet('Item_Saran');
  const id = 'IT-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  sh.appendRow([id, body.idKunjungan, body.namaKomponen, body.qty, body.keterangan || '',
    '', '', body.ketersediaanPart || '', '', body.fotoUrl || '', new Date(), '', body.hargaSatuan || '']);
  return { success: true, idItem: id };
}

/** Teknisi tambah 1 item jasa (Nama Jasa) ke kunjungan yang sudah ada */
function addItemJasa(body) {
  if (!body.namaJasa) {
    return { error: 'Nama jasa wajib diisi' };
  }
  const sh = getSheet('Item_Jasa');
  const id = 'JS-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  sh.appendRow([id, body.idKunjungan, body.namaJasa, body.waktu || '', body.hargaSatuan || '',
    body.keterangan || '', new Date()]);
  return { success: true, idJasa: id };
}

/** Teknisi tambah 1 baris SSC ke kunjungan yang sudah ada */
function addItemSSC(body) {
  if (!body.ssc) {
    return { error: 'SSC wajib diisi' };
  }
  const sh = getSheet('Item_SSC');
  const id = 'SSC-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  sh.appendRow([id, body.idKunjungan, body.ssc, body.status || '', body.alasan || '', new Date()]);
  return { success: true, idSSC: id };
}

/** Teknisi tambah 1 baris Technical Information ke kunjungan yang sudah ada */
function addItemTechnicalInfo(body) {
  if (!body.technicalInformation) {
    return { error: 'Technical Information wajib diisi' };
  }
  const sh = getSheet('Item_TechnicalInfo');
  const id = 'TI-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  sh.appendRow([id, body.idKunjungan, body.technicalInformation, body.status || '', body.alasan || '', new Date()]);
  return { success: true, idTI: id };
}

/** ============ PARTMAN / FOREMAN: lengkapi data part ============ */
function updatePartmanItem(body) {
  const sh = getSheet('Item_Saran');
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === body.idItem) {
      if (body.nomorPart !== undefined) sh.getRange(r + 1, 6).setValue(body.nomorPart);
      if (body.estimasiHarga !== undefined) sh.getRange(r + 1, 7).setValue(body.estimasiHarga);
      if (body.ketersediaanPart !== undefined) sh.getRange(r + 1, 8).setValue(body.ketersediaanPart);
      if (body.keteranganPartman !== undefined) sh.getRange(r + 1, 9).setValue(body.keteranganPartman);
      sh.getRange(r + 1, 12).setValue(new Date());
      return { success: true };
    }
  }
  return { error: 'Item tidak ditemukan' };
}

/** Upload foto komponen rusak (base64) ke Drive, kembalikan URL */
function uploadFoto(body) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const bytes = Utilities.base64Decode(body.base64Data);
  const blob = Utilities.newBlob(bytes, body.mimeType || 'image/jpeg', body.fileName || ('foto_' + new Date().getTime() + '.jpg'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { success: true, url: file.getUrl(), fileId: file.getId() };
}

/** ============ MASTER DATA (isi dropdown) ============ */
function getMasterData() {
  const sh = getSheet('MasterData');
  const headers = HEADERS.MasterData;
  const lastRow = sh.getLastRow();
  const result = {};
  headers.forEach(h => result[h] = []);
  if (lastRow < 2) return result;

  const values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  headers.forEach((h, col) => {
    values.forEach(row => {
      const v = row[col];
      if (v !== '' && v !== null && v !== undefined) result[h].push(v);
    });
  });
  return result;
}

/** ============ LIST & DETAIL ============ */
function listKunjungan(params) {
  const sh = getSheet('Kunjungan');
  const data = sh.getDataRange().getValues();
  const header = data.shift();
  let rows = data.map(r => rowToObj(header, r));
  if (params.sa) rows = rows.filter(r => r.SA === params.sa);
  if (params.status) rows = rows.filter(r => r.Status === params.status);
  return rows;
}

function getKunjunganDetail(id) {
  const kj = listKunjungan({}).find(r => r.ID_Kunjungan === id);
  if (!kj) return { error: 'Tidak ditemukan' };
  const shItem = getSheet('Item_Saran');
  const data = shItem.getDataRange().getValues();
  const header = data.shift();
  const items = data.map(r => rowToObj(header, r)).filter(it => it.ID_Kunjungan === id);

  const shJasa = getSheet('Item_Jasa');
  const dataJasa = shJasa.getDataRange().getValues();
  const headerJasa = dataJasa.shift();
  const itemsJasa = dataJasa.map(r => rowToObj(headerJasa, r)).filter(it => it.ID_Kunjungan === id);

  const shSSC = getSheet('Item_SSC');
  const dataSSC = shSSC.getDataRange().getValues();
  const headerSSC = dataSSC.shift();
  const itemsSSC = dataSSC.map(r => rowToObj(headerSSC, r)).filter(it => it.ID_Kunjungan === id);

  const shTI = getSheet('Item_TechnicalInfo');
  const dataTI = shTI.getDataRange().getValues();
  const headerTI = dataTI.shift();
  const itemsTechnicalInfo = dataTI.map(r => rowToObj(headerTI, r)).filter(it => it.ID_Kunjungan === id);

  return { kunjungan: kj, items: items, itemsJasa: itemsJasa, itemsSSC: itemsSSC, itemsTechnicalInfo: itemsTechnicalInfo };
}

function rowToObj(header, row) {
  const obj = {};
  header.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

/** Sisipkan baris data tepat di posisi `placeholder` (mis. "{{TABEL_ITEM}}").
 *  `rowsWithHeader[0]` adalah header, sisanya baris data. Ada 2 kasus:
 *  1. Placeholder ada di dalam sel tabel yang sudah digambar di template
 *     (header, baris kosong berisi placeholder, mungkin ada baris Total di
 *     bawahnya) - baris data disisipkan LANGSUNG ke tabel itu di posisi
 *     baris placeholder, lalu baris placeholder aslinya dibuang. Header dan
 *     baris-baris lain (mis. baris Total) tetap utuh, dan lebar kolom
 *     mengikuti desain asli template - tidak perlu diatur manual.
 *  2. Placeholder ada di paragraf biasa (bukan di dalam tabel) - tabel baru
 *     (lengkap dengan header) disisipkan di situ, paragrafnya dikosongkan
 *     (bukan dihapus - Google Docs tidak mengizinkan body section diakhiri
 *     tabel).
 *  Jika placeholder tidak ditemukan sama sekali, tabel baru ditambahkan di
 *  akhir dokumen sebagai fallback. */
function insertTableAtPlaceholder(body, placeholder, rowsWithHeader) {
  const found = body.findText(placeholder);
  if (!found) return body.appendTable(rowsWithHeader);

  let el = found.getElement();
  while (el.getParent() && el.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
    el = el.getParent();
  }

  if (el.getType() === DocumentApp.ElementType.TABLE) {
    return insertRowsIntoExistingTable(el, placeholder, rowsWithHeader.slice(1));
  }

  const index = body.getChildIndex(el);
  const table = body.insertTable(index, rowsWithHeader);
  el.editAsText().setText('');
  return table;
}

/** Cari row di `table` yang mengandung `placeholder`, ganti dengan `dataRows`
 *  (tanpa header - tabel lama sudah punya headernya sendiri), dengan jumlah
 *  kolom mengikuti row placeholder aslinya (bukan jumlah kolom dataRows). */
function insertRowsIntoExistingTable(table, placeholder, dataRows) {
  let rowIndex = -1;
  for (let r = 0; r < table.getNumRows(); r++) {
    if (table.getRow(r).findText(placeholder)) { rowIndex = r; break; }
  }
  if (rowIndex === -1) rowIndex = table.getNumRows() - 1;
  const numCols = table.getRow(rowIndex).getNumCells();

  if (!dataRows.length) {
    for (let c = 0; c < numCols; c++) table.getRow(rowIndex).getCell(c).setText('');
    return table;
  }

  dataRows.forEach((rowValues, i) => {
    const newRow = table.insertTableRow(rowIndex + i);
    for (let c = 0; c < numCols; c++) {
      const cell = newRow.appendTableCell(rowValues[c] !== undefined ? String(rowValues[c]) : '');
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
  });
  table.removeRow(rowIndex + dataRows.length); // baris placeholder asli, sudah bergeser ke bawah
  return table;
}

function formatRupiahOrDash(v) {
  const n = Number(v) || 0;
  return n ? 'Rp ' + n.toLocaleString('id-ID') : '-';
}

const PHOTOS_PER_ROW = 3;

/** Lampirkan beberapa foto (dari URL Drive share-link) di akhir dokumen dalam
 *  grid beberapa kolom sejajar (caption "Foto: Label" di atas tiap gambar),
 *  bukan 1 foto per halaman penuh - supaya tidak makan tempat berlebihan.
 *  `photos` = [{label, fotoUrl}, ...]; entri tanpa fotoUrl diabaikan. */
function appendPhotoGrid(body, photos) {
  const valid = photos.filter(p => p.fotoUrl);
  if (!valid.length) return;

  const contentWidth = body.getPageWidth() - body.getMarginLeft() - body.getMarginRight();

  for (let i = 0; i < valid.length; i += PHOTOS_PER_ROW) {
    const rowPhotos = valid.slice(i, i + PHOTOS_PER_ROW);
    const table = body.appendTable([rowPhotos.map(() => '')]);
    table.setBorderWidth(0);
    const cellWidth = contentWidth / rowPhotos.length;
    rowPhotos.forEach((p, c) => {
      const cell = table.getCell(0, c);
      const captionPara = cell.getChild(0).asParagraph();
      captionPara.setText('Foto: ' + p.label);
      captionPara.setItalic(true);
      try {
        const fileId = p.fotoUrl.match(/[-\w]{25,}/)[0];
        const imgBlob = DriveApp.getFileById(fileId).getBlob();
        const img = cell.appendImage(imgBlob);
        const naturalWidth = img.getWidth();
        const naturalHeight = img.getHeight();
        const targetWidth = cellWidth - 20;
        const targetHeight = naturalHeight * (targetWidth / naturalWidth);
        img.setWidth(targetWidth);
        img.setHeight(targetHeight);
      } catch (e) { /* skip jika gagal ambil gambar */ }
    });
  }
}

/** Generate gambar QR code (stempel info, bukan link verifikasi) lewat API publik
 *  api.qrserver.com - hasil scan cuma menampilkan teks yang dikirim. */
function generateQrBlob(text) {
  const url = 'https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=' + encodeURIComponent(text);
  const resp = UrlFetchApp.fetch(url);
  return resp.getBlob().setName('qr_signing.png');
}

/** Cari placeholder teks di sebuah section (Body/HeaderSection/FooterSection, termasuk
 *  di dalam sel tabel), kosongkan paragrafnya, lalu sisipkan gambar di situ. */
function insertImageAtPlaceholder(section, placeholder, imageBlob, widthPt) {
  const found = section.findText(placeholder);
  if (!found) return false;
  const para = found.getElement().getParent().asParagraph();
  para.setText(' '); // Docs tidak izinkan paragraf dikosongkan total via setText('')
  const img = para.appendInlineImage(imageBlob);
  if (widthPt) {
    const ratio = img.getHeight() / img.getWidth();
    img.setWidth(widthPt);
    img.setHeight(widthPt * ratio);
  }
  return true;
}

/** Placeholder bisa ditaruh user di body, header, ATAU footer dokumen (footer/header
 *  adalah section terpisah dari body di Google Docs) - coba replace/insert di ketiganya. */
function replaceTextEverywhere(doc, placeholder, value) {
  [doc.getBody(), doc.getHeader(), doc.getFooter()].forEach(section => {
    if (section) section.replaceText(placeholder, value);
  });
}
function insertImageAtPlaceholderEverywhere(doc, placeholder, imageBlob, widthPt) {
  const sections = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(s => s);
  for (const section of sections) {
    if (insertImageAtPlaceholder(section, placeholder, imageBlob, widthPt)) return true;
  }
  return false;
}

/** ============ GENERATE PDF REPORT (Teknisi draft / Foreman -> SA) ============
 *  printedBy: 'teknisi' (cetak draft langsung setelah input, status TIDAK berubah,
 *  belum ada tanda tangan/QR) atau 'foreman' (default, laporan final: status jadi
 *  "Report Terkirim", QR signing dibuat, PDF_URL & Nama_Foreman disimpan). */
function generateReport(idKunjungan, namaForeman, printedBy) {
  printedBy = printedBy === 'teknisi' ? 'teknisi' : 'foreman';
  const detail = getKunjunganDetail(idKunjungan);
  if (detail.error) return detail;
  const kj = detail.kunjungan;
  const items = detail.items || [];
  const itemsJasa = detail.itemsJasa || [];
  const itemsSSC = detail.itemsSSC || [];

  // Duplikat template Google Docs, isi placeholder, export ke PDF
  const templateFile = DriveApp.getFileById(PDF_TEMPLATE_ID);
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const fileSuffix = printedBy === 'teknisi' ? '_draft' : '';
  const docCopy = templateFile.makeCopy('Report_' + kj.No_Polisi + '_' + idKunjungan + fileSuffix, folder);
  const doc = DocumentApp.openById(docCopy.getId());
  const body = doc.getBody();

  // Header
  body.replaceText('{{NO_POLISI}}', kj.No_Polisi || '');
  body.replaceText('{{NAMA_CUSTOMER}}', kj.Nama_Customer || '');
  body.replaceText('{{TIPE_MOBIL}}', kj.Tipe_Mobil || '');
  body.replaceText('{{TIPE_SERVICE}}', kj.Tipe_Service || '');
  body.replaceText('{{TANGGAL_SERVICE}}', kj.Tanggal_Masuk ? Utilities.formatDate(new Date(kj.Tanggal_Masuk), 'GMT+7', 'dd/MM/yyyy') : '-');
  body.replaceText('{{SERVICE_ADVISOR}}', kj.SA || '');
  body.replaceText('{{MEKANIK}}', kj.Teknisi || '');

  // Proses Service
  body.replaceText('{{PEKERJAAN}}', kj.Pekerjaan || '-');
  body.replaceText('{{REQUEST_CUSTOMER}}', kj.Request_Customer || '-');
  body.replaceText('{{AKI_STATUS}}', kj.Aki_Status || '-');
  body.replaceText('{{AKI_KETERANGAN}}', kj.Aki_Keterangan || '-');
  body.replaceText('{{AKI_HARGA}}', formatRupiahOrDash(kj.Aki_Harga));
  body.replaceText('{{BAN_STATUS}}', kj.Ban_Status || '-');
  body.replaceText('{{BAN_KETERANGAN}}', kj.Ban_Keterangan || '-');
  body.replaceText('{{BAN_MERK1}}', kj.Ban_Merk1 || '-');
  body.replaceText('{{BAN_HARGA1}}', formatRupiahOrDash(kj.Ban_Harga1));
  body.replaceText('{{BAN_MERK2}}', kj.Ban_Merk2 || '-');
  body.replaceText('{{BAN_HARGA2}}', formatRupiahOrDash(kj.Ban_Harga2));
  body.replaceText('{{KAMPASREM_STATUS}}', kj.KampasRem_Status || '-');
  body.replaceText('{{KAMPASREM_KETERANGAN}}', kj.KampasRem_Keterangan || '-');
  body.replaceText('{{KAMPASREM_HARGA}}', formatRupiahOrDash(kj.KampasRem_Harga));
  body.replaceText('{{WIPER_STATUS}}', kj.Wiper_Status || '-');
  body.replaceText('{{WIPER_KETERANGAN}}', kj.Wiper_Keterangan || '-');
  body.replaceText('{{WIPER_HARGA}}', formatRupiahOrDash(kj.Wiper_Harga));
  body.replaceText('{{LAMPU_STATUS}}', kj.Lampu_Status || '-');
  body.replaceText('{{LAMPU_KETERANGAN}}', kj.Lampu_Keterangan || '-');
  body.replaceText('{{LAMPU_HARGA}}', formatRupiahOrDash(kj.Lampu_Harga));

  // Saran Perbaikan - Nama Jasa
  let totalJasa = 0;
  const tableJasa = [['Nama Jasa', 'Waktu', 'Harga Satuan', 'Keterangan']];
  itemsJasa.forEach(it => {
    const harga = Number(it.Harga_Satuan) || 0;
    totalJasa += harga;
    tableJasa.push([it.Nama_Jasa, it.Waktu || '-', formatRupiahOrDash(harga), it.Keterangan || '-']);
  });
  insertTableAtPlaceholder(body, '{{TABEL_JASA}}', tableJasa);
  body.replaceText('{{TOTAL_JASA}}', 'Rp ' + totalJasa.toLocaleString('id-ID'));

  // Saran Perbaikan - Nama Parts (harga pakai Estimasi_Harga Foreman, fallback ke estimasi Teknisi)
  let totalParts = 0;
  const tableParts = [['Nama Parts', 'Qty', 'Harga Satuan', 'Total', 'Keterangan']];
  items.forEach(it => {
    const qty = Number(it.Qty) || 1;
    const harga = Number(it.Estimasi_Harga) || Number(it.Harga_Satuan_Teknisi) || 0;
    const rowTotal = harga * qty;
    totalParts += rowTotal;
    tableParts.push([it.Nama_Komponen, qty, formatRupiahOrDash(harga),
      formatRupiahOrDash(rowTotal), it.Keterangan_Partman || it.Keterangan_Teknisi || '-']);
  });
  insertTableAtPlaceholder(body, '{{TABEL_PARTS}}', tableParts);
  body.replaceText('{{TOTAL_PARTS}}', 'Rp ' + totalParts.toLocaleString('id-ID'));
  body.replaceText('{{TOTAL_ESTIMASI}}', 'Rp ' + (totalJasa + totalParts).toLocaleString('id-ID'));

  // Hybrid Health Check
  body.replaceText('{{MOBIL_HYBRID}}', kj.Mobil_Hybrid || 'No');
  body.replaceText('{{SBE_50K_100K}}', kj.SBE_50K_100K || '-');
  body.replaceText('{{MOBIL_2_5_TAHUN}}', kj.Mobil_2_5_Tahun || '-');
  body.replaceText('{{PENGERJAAN_HHC}}', kj.HHC || '-');
  body.replaceText('{{HASIL_HHC}}', kj.HHC_Hasil || '-');

  // Special Service Campaign
  const tableSSC = [['Jenis SSC', 'Status', 'Alasan']];
  itemsSSC.forEach(it => tableSSC.push([it.SSC, it.Status || '-', it.Alasan || '-']));
  insertTableAtPlaceholder(body, '{{TABEL_SSC}}', tableSSC);

  // Uji Emisi
  body.replaceText('{{UJI_EMISI_STATUS}}', kj.UjiEmisi_Status || 'Tidak Lulus/Belum Uji Emisi/Kadaluarsa');

  // Tanda tangan - hanya terisi saat laporan final dari Foreman
  const now = new Date();
  body.replaceText('{{TANGGAL}}', printedBy === 'foreman' ? Utilities.formatDate(now, 'GMT+7', 'dd/MM/yyyy') : '-');
  body.replaceText('{{NAMA_FOREMAN}}', printedBy === 'foreman' ? (namaForeman || '') : '-');

  // QR signing (stempel info, bukan link verifikasi) - hanya saat Foreman finalisasi
  let qrWarning = null;
  if (printedBy === 'foreman') {
    const qrText = 'LAPORAN DITANDATANGANI\nForeman: ' + (namaForeman || '-') +
      '\nNo. Polisi: ' + (kj.No_Polisi || '-') +
      '\nID Kunjungan: ' + idKunjungan +
      '\nTanggal: ' + Utilities.formatDate(now, 'GMT+7', 'dd/MM/yyyy HH:mm');
    try {
      const qrBlob = generateQrBlob(qrText);
      const inserted = insertImageAtPlaceholderEverywhere(doc, '{{QR_SIGNING}}', qrBlob, 80);
      if (!inserted) qrWarning = 'Placeholder {{QR_SIGNING}} tidak ditemukan di template';
    } catch (e) {
      qrWarning = 'Gagal generate/insert QR: ' + e.message;
      replaceTextEverywhere(doc, '{{QR_SIGNING}}', '');
    }
  } else {
    replaceTextEverywhere(doc, '{{QR_SIGNING}}', '');
  }

  // Info cetak - footer kecil khusus versi draft Teknisi
  replaceTextEverywhere(doc, '{{PRINT_INFO}}', printedBy === 'teknisi'
    ? ('Print by Teknisi - ' + Utilities.formatDate(now, 'GMT+7', 'dd/MM/yyyy HH:mm'))
    : '');

  // Dokumentasi: foto part yang ada fotonya + foto sertifikat uji emisi
  const photos = items.map(it => ({ label: it.Nama_Komponen, fotoUrl: it.Foto_URL }));
  photos.push({ label: 'Sertifikat Uji Emisi', fotoUrl: kj.UjiEmisi_Foto_URL });
  appendPhotoGrid(body, photos);

  doc.saveAndClose();

  const pdfBlob = docCopy.getAs('application/pdf');
  const pdfFile = folder.createFile(pdfBlob).setName('Report_' + kj.No_Polisi + '_' + idKunjungan + fileSuffix + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getFileById(docCopy.getId()).setTrashed(true); // hapus file docs sementara, sisakan PDF

  // update status kunjungan + simpan nama Foreman yang mengirim - hanya untuk laporan final
  // (cetak draft oleh Teknisi tidak mengubah status, Foreman tetap wajib QC/finalisasi)
  if (printedBy === 'foreman') {
    const sh = getSheet('Kunjungan');
    const data = sh.getDataRange().getValues();
    const header = data[0];
    const statusCol = header.indexOf('Status');
    const pdfCol = header.indexOf('PDF_URL');
    const foremanCol = header.indexOf('Nama_Foreman');
    const perluFollowUpCol = header.indexOf('Perlu_FollowUp');
    // Kalau Teknisi/Foreman tidak menyarankan Jasa maupun Parts sama sekali,
    // kunjungan ini tidak perlu masuk alur follow up SA.
    const perluFollowUp = (itemsJasa.length + items.length) > 0 ? 'Ya' : 'Tidak';
    for (let r = 1; r < data.length; r++) {
      if (data[r][0] === idKunjungan) {
        sh.getRange(r + 1, statusCol + 1).setValue('Report Terkirim');
        sh.getRange(r + 1, pdfCol + 1).setValue(pdfFile.getUrl());
        if (foremanCol !== -1 && namaForeman) sh.getRange(r + 1, foremanCol + 1).setValue(namaForeman);
        if (perluFollowUpCol !== -1) sh.getRange(r + 1, perluFollowUpCol + 1).setValue(perluFollowUp);
        break;
      }
    }
  }
  return { success: true, pdfUrl: pdfFile.getUrl(), totalEstimasi: totalJasa + totalParts, qrWarning: qrWarning };
}

/** ============ SA: FOLLOW UP (H+3) ============ */
function updateFollowUp(body) {
  const sh = getSheet('Kunjungan');
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === body.idKunjungan) {
      sh.getRange(r + 1, 12).setValue(body.statusFollowUp);   // kolom L
      sh.getRange(r + 1, 13).setValue(body.catatan || '');    // kolom M
      sh.getRange(r + 1, 14).setValue(new Date());             // kolom N
      sh.getRange(r + 1, 9).setValue('Follow Up');              // kolom I Status
      return { success: true };
    }
  }
  return { error: 'Kunjungan tidak ditemukan' };
}

/** ============ DASHBOARD KEPALA BENGKEL ============ */
function getDashboard(params) {
  const rows = listKunjungan({});
  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfWeek = new Date(startOfDay.getTime() - startOfDay.getDay() * 86400000);
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  // Hitung jumlah Item_Saran per kunjungan SEKALI di sini (bukan getKunjunganDetail per
  // baris seperti sebelumnya, yang re-read seluruh sheet Item_* untuk TIAP kunjungan -
  // O(n) pemanggilan penuh, sangat lambat begitu data bertambah banyak).
  const itemSh = getSheet('Item_Saran');
  const itemData = itemSh.getDataRange().getValues();
  const itemHeader = itemData[0];
  const itemKunjunganCol = itemHeader.indexOf('ID_Kunjungan');
  const itemCountMap = {};
  for (let i = 1; i < itemData.length; i++) {
    const kid = itemData[i][itemKunjunganCol];
    if (!kid) continue;
    itemCountMap[kid] = (itemCountMap[kid] || 0) + 1;
  }

  const bySA = {};
  let totalCustomerHariIni = 0, totalCustomerMingguIni = 0, totalCustomerBulanIni = 0;

  rows.forEach(r => {
    const tgl = new Date(r.Tanggal_Masuk);
    const sa = r.SA || '(Tanpa SA)';
    if (!bySA[sa]) bySA[sa] = { hari: 0, minggu: 0, bulan: 0, customerHari: 0, customerMinggu: 0, customerBulan: 0, deal: 0, tidakDeal: 0 };

    // hitung jumlah item saran per kunjungan untuk agregasi "berapa yang disarankan"
    const jumlahItem = itemCountMap[r.ID_Kunjungan] || 0;

    if (tgl >= startOfDay) { bySA[sa].hari += jumlahItem; bySA[sa].customerHari++; totalCustomerHariIni++; }
    if (tgl >= startOfWeek) { bySA[sa].minggu += jumlahItem; bySA[sa].customerMinggu++; totalCustomerMingguIni++; }
    if (tgl >= startOfMonth) { bySA[sa].bulan += jumlahItem; bySA[sa].customerBulan++; totalCustomerBulanIni++; }

    if (r.Status_FollowUp === 'Deal - Belum Datang' || r.Status_FollowUp === 'Deal - Sudah Datang') bySA[sa].deal++;
    if (r.Status_FollowUp === 'Tidak Deal') bySA[sa].tidakDeal++;
  });

  return {
    perSA: bySA,
    totalCustomer: { hariIni: totalCustomerHariIni, mingguIni: totalCustomerMingguIni, bulanIni: totalCustomerBulanIni },
    totalKunjungan: rows.length
  };
}
