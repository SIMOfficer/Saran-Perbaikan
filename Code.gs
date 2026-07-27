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
 * I: Status  (Menunggu Partman / Siap Report / Report Terkirim / Follow Up / Deal / Tidak Deal)
 * J: PDF_URL
 * K: Tanggal_FollowUp_Rencana  (otomatis = Tanggal_Masuk + 3 hari)
 * L: Status_FollowUp  (Belum / Deal / Tidak Deal / Reschedule)
 * M: Catatan_FollowUp
 * N: Tanggal_FollowUp_Aktual
 *
 * SHEET "Item_Saran" (1 baris = 1 komponen yang disarankan)
 * A: ID_Item
 * B: ID_Kunjungan   (FK ke sheet Kunjungan)
 * C: Nama_Komponen        (wajib - Teknisi)
 * D: Qty                  (wajib - Teknisi)
 * E: Keterangan_Teknisi   (opsional - Teknisi)
 * F: Nomor_Part           (diisi Partman/Foreman)
 * G: Estimasi_Harga       (diisi Partman/Foreman)
 * H: Ketersediaan_Part    (diisi Partman/Foreman)
 * I: Keterangan_Partman   (opsional - Partman/Foreman)
 * J: Foto_URL             (opsional, link Drive)
 * K: Diisi_Teknisi_At
 * L: Diisi_Partman_At
 * ============================================================
 */

const SPREADSHEET_ID = '1_JYeu0uYI1CxLA2Y5EMFDFNFaCZnhibG_--o-YGRRqA'; // ID Google Sheet (database)
const DRIVE_FOLDER_ID = '1A6VLdeox-bhZGS-u9XsyfOnOfTF41viz'; // Folder khusus foto komponen
const PDF_TEMPLATE_ID = '1IP_2tMxMMXprOvNy9HbsTBrS4LK2lw6cDfV2UThQ1VQ'; // Template dokumen report
const CODE_VERSION = 'v9-technical-information'; // Ganti tiap perubahan, dipakai action=version untuk cek deployment

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
    'SSC_Terlibat'],
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
    'Technical_Information','Alasan_Technical']
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
      case 'addItemSaran': result = addItemSaran(body); break;
      case 'addItemJasa': result = addItemJasa(body); break;
      case 'addItemSSC': result = addItemSSC(body); break;
      case 'addItemTechnicalInfo': result = addItemTechnicalInfo(body); break;
      case 'updatePartman': result = updatePartmanItem(body); break;
      case 'uploadFoto': result = uploadFoto(body); break;
      case 'generateReport': result = generateReport(body.idKunjungan); break;
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
  sh.appendRow([id, now, body.tanggalMasuk || now, body.noPolisi, body.namaCustomer,
    body.noHp || '', body.sa, body.teknisi, 'Menunggu Partman', '', followUpDate, 'Belum', '', '',
    body.tipeMobil || '', body.tipeService || '',
    p.pekerjaan || '', p.requestCustomer || '',
    aki.status || 'OK', aki.keterangan || '', aki.harga || '',
    ban.status || 'OK', ban.keterangan || '', ban.merk1 || '', ban.harga1 || '', ban.merk2 || '', ban.harga2 || '',
    kampasRem.status || 'OK', kampasRem.keterangan || '', kampasRem.harga || '',
    wiper.status || 'OK', wiper.keterangan || '', wiper.harga || '',
    lampu.status || 'OK', lampu.keterangan || '', lampu.harga || '',
    hhc.mobilHybrid || 'No', hhc.sbe50k100k || '', hhc.mobil2to5Tahun || '', hhc.hhc || '', hhc.hasil || '',
    body.sscTerlibat || 'No']);

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

/** Teknisi tambah 1 item part (Nama Parts) ke kunjungan yang sudah ada */
function addItemSaran(body) {
  if (!body.namaKomponen || !body.qty) {
    return { error: 'Nama komponen dan Qty wajib diisi' };
  }
  const sh = getSheet('Item_Saran');
  const id = 'IT-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  sh.appendRow([id, body.idKunjungan, body.namaKomponen, body.qty, body.keterangan || '',
    '', '', '', '', body.fotoUrl || '', new Date(), '', body.hargaSatuan || '']);
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

/** Sisipkan tabel tepat sebelum paragraf yang mengandung `placeholder`
 *  (mis. "{{TABEL_ITEM}}"), bukan selalu di akhir dokumen. Teks paragraf
 *  placeholder dikosongkan (bukan dihapus elemennya) dan disisakan sebagai
 *  baris kosong setelah tabel - Google Docs tidak mengizinkan sebuah body
 *  section diakhiri oleh tabel, jadi paragraf itu wajib tetap ada kalau
 *  placeholder-nya berada di baris terakhir dokumen. Jika placeholder tidak
 *  ditemukan, tabel ditambahkan di akhir dokumen sebagai fallback. */
function insertTableAtPlaceholder(body, placeholder, tableData) {
  const found = body.findText(placeholder);
  if (!found) return body.appendTable(tableData);

  let el = found.getElement();
  while (el.getParent() && el.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
    el = el.getParent();
  }
  const index = body.getChildIndex(el);
  const table = body.insertTable(index, tableData);
  el.editAsText().setText('');
  return table;
}

/** ============ GENERATE PDF REPORT UNTUK CUSTOMER ============ */
function generateReport(idKunjungan) {
  const detail = getKunjunganDetail(idKunjungan);
  if (detail.error) return detail;
  const kj = detail.kunjungan;
  const items = detail.items;

  // Duplikat template Google Docs, isi placeholder, export ke PDF
  const templateFile = DriveApp.getFileById(PDF_TEMPLATE_ID);
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const docCopy = templateFile.makeCopy('Report_' + kj.No_Polisi + '_' + idKunjungan, folder);
  const doc = DocumentApp.openById(docCopy.getId());
  const body = doc.getBody();

  body.replaceText('{{NO_POLISI}}', kj.No_Polisi);
  body.replaceText('{{NAMA_CUSTOMER}}', kj.Nama_Customer);
  body.replaceText('{{TANGGAL}}', Utilities.formatDate(new Date(kj.Tanggal_Masuk), 'GMT+7', 'dd/MM/yyyy'));
  body.replaceText('{{SA}}', kj.SA);
  body.replaceText('{{TEKNISI}}', kj.Teknisi);

  let total = 0;
  const tableData = [['No', 'Komponen', 'Qty', 'No. Part', 'Est. Harga', 'Ketersediaan', 'Ket.']];
  items.forEach((it, i) => {
    const harga = Number(it.Estimasi_Harga) || 0;
    total += harga * Number(it.Qty || 1);
    tableData.push([i + 1, it.Nama_Komponen, it.Qty, it.Nomor_Part || '-',
      harga ? harga.toLocaleString('id-ID') : '-', it.Ketersediaan_Part || '-',
      it.Keterangan_Partman || it.Keterangan_Teknisi || '-']);
  });
  // Sisipkan tabel tepat di posisi placeholder {{TABEL_ITEM}} (fallback: akhir dokumen)
  const table = insertTableAtPlaceholder(body, '{{TABEL_ITEM}}', tableData);
  body.replaceText('{{TOTAL_ESTIMASI}}', 'Rp ' + total.toLocaleString('id-ID'));

  // Lampirkan foto komponen (jika ada) di halaman berikut
  items.forEach(it => {
    if (it.Foto_URL) {
      body.appendParagraph('Foto: ' + it.Nama_Komponen);
      try {
        const fileId = it.Foto_URL.match(/[-\w]{25,}/)[0];
        const imgBlob = DriveApp.getFileById(fileId).getBlob();
        body.appendImage(imgBlob).setWidth(300);
      } catch (e) { /* skip jika gagal ambil gambar */ }
    }
  });

  doc.saveAndClose();

  const pdfBlob = docCopy.getAs('application/pdf');
  const pdfFile = folder.createFile(pdfBlob).setName('Report_' + kj.No_Polisi + '_' + idKunjungan + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getFileById(docCopy.getId()).setTrashed(true); // hapus file docs sementara, sisakan PDF

  // update status kunjungan
  const sh = getSheet('Kunjungan');
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === idKunjungan) {
      sh.getRange(r + 1, 9).setValue('Report Terkirim'); // kolom I = Status
      sh.getRange(r + 1, 10).setValue(pdfFile.getUrl()); // kolom J = PDF_URL
      break;
    }
  }
  return { success: true, pdfUrl: pdfFile.getUrl(), totalEstimasi: total };
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

  const bySA = {};
  let totalCustomerHariIni = 0, totalCustomerMingguIni = 0, totalCustomerBulanIni = 0;

  rows.forEach(r => {
    const tgl = new Date(r.Tanggal_Masuk);
    const sa = r.SA || '(Tanpa SA)';
    if (!bySA[sa]) bySA[sa] = { hari: 0, minggu: 0, bulan: 0, customerHari: 0, customerMinggu: 0, customerBulan: 0, deal: 0, tidakDeal: 0 };

    // hitung jumlah item saran per kunjungan untuk agregasi "berapa yang disarankan"
    const detail = getKunjunganDetail(r.ID_Kunjungan);
    const jumlahItem = detail.items ? detail.items.length : 0;

    if (tgl >= startOfDay) { bySA[sa].hari += jumlahItem; bySA[sa].customerHari++; totalCustomerHariIni++; }
    if (tgl >= startOfWeek) { bySA[sa].minggu += jumlahItem; bySA[sa].customerMinggu++; totalCustomerMingguIni++; }
    if (tgl >= startOfMonth) { bySA[sa].bulan += jumlahItem; bySA[sa].customerBulan++; totalCustomerBulanIni++; }

    if (r.Status_FollowUp === 'Deal') bySA[sa].deal++;
    if (r.Status_FollowUp === 'Tidak Deal') bySA[sa].tidakDeal++;
  });

  return {
    perSA: bySA,
    totalCustomer: { hariIni: totalCustomerHariIni, mingguIni: totalCustomerMingguIni, bulanIni: totalCustomerBulanIni },
    totalKunjungan: rows.length
  };
}
