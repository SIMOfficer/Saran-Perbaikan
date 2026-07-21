# Sistem Manajemen Saran Perbaikan - Prototype

## Arsitektur
- **Frontend**: `index.html` (single page, vanilla JS) → di-host di GitHub Pages
- **Backend**: `Code.gs` (Google Apps Script) → berfungsi sebagai API, jadi tidak perlu server sendiri
- **Database**: Google Spreadsheet (2 sheet: `Kunjungan` dan `Item_Saran`)
- **Storage foto**: Google Drive folder

Alur data: `Browser (GitHub Pages) → Google Apps Script Web App (API) → Google Sheets/Drive`

## Langkah Setup

### 1. Siapkan Google Spreadsheet
- Buat 1 file Google Sheet baru sebagai database, salin ID-nya dari URL
  (`https://docs.google.com/spreadsheets/d/**ID_INI**/edit`)

### 2. Siapkan folder Google Drive untuk foto
- Buat folder khusus, salin ID folder-nya dari URL

### 3. Siapkan template Google Docs untuk PDF report
- Buat 1 Google Docs berisi kop/letterhead perusahaan + placeholder teks:
  `{{NO_POLISI}}`, `{{NAMA_CUSTOMER}}`, `{{TANGGAL}}`, `{{SA}}`, `{{TEKNISI}}`, `{{TOTAL_ESTIMASI}}`
- Salin ID Docs-nya

### 4. Deploy Apps Script
1. Buka [script.google.com](https://script.google.com) → New Project
2. Copy-paste isi `Code.gs`
3. Isi 3 variabel di baris atas: `SPREADSHEET_ID`, `DRIVE_FOLDER_ID`, `PDF_TEMPLATE_ID`
4. Jalankan fungsi `setupSheets` sekali (untuk buat header sheet otomatis) — akan diminta izin akses, setujui
5. Deploy → New deployment → Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone** (atau **Anyone within [organisasi]** kalau pakai Google Workspace)
6. Salin URL Web App yang muncul (formatnya `https://script.google.com/macros/s/xxxx/exec`)

### 5. Hubungkan Frontend
- Buka `index.html`, ganti baris:
  ```js
  const API_URL = "PASTE_URL_WEB_APP_APPS_SCRIPT_DI_SINI";
  ```
  dengan URL dari langkah 4.6
- Push ke repo GitHub, aktifkan GitHub Pages (Settings → Pages → source: main branch)

## Alur Kerja Aplikasi
1. **Teknisi** input temuan saat servis → wajib isi Nama Komponen + Qty (+ Keterangan opsional) → status `Menunggu Partman`
2. **Partman/Foreman** melengkapi Nomor Part, Estimasi Harga, Ketersediaan Part (+ keterangan tambahan) → klik "Selesai & Buat Report PDF" → sistem generate PDF A4 (termasuk total biaya & foto) → status `Report Terkirim`
3. Sistem otomatis menghitung **tanggal follow up H+3** saat kunjungan dibuat
4. **Service Advisor** melakukan follow up pada tanggal tsb, input hasil (Deal / Tidak Deal / Reschedule)
5. **Kepala Bengkel** memantau dashboard: total saran per SA (harian/mingguan/bulanan) + jumlah customer yang datang

## Catatan Prototype
- Belum ada sistem login/otentikasi sungguhan — role dipilih manual via dropdown di kanan atas (cocok untuk uji coba internal, sebelum nanti ditambah Google Sign-In / OAuth kalau mau produksi)
- Upload foto di frontend belum disambungkan ke `uploadFoto` endpoint — perlu ditambahkan konversi file ke base64 sebelum dikirim (rangka fungsinya sudah ada di `Code.gs`)
- Placeholder tabel item di template Docs (`{{TABEL_ITEM}}`) harus berada sendirian di barisnya sendiri — `generateReport()` akan mengganti seluruh paragraf tersebut dengan tabel item. Jika placeholder tidak ditemukan, tabel otomatis ditambahkan di akhir dokumen sebagai fallback

## Struktur Data

**Sheet Kunjungan** (1 baris = 1 unit yang masuk):
ID_Kunjungan, Timestamp, Tanggal_Masuk, No_Polisi, Nama_Customer, No_HP_Customer, SA, Teknisi, Status, PDF_URL, Tanggal_FollowUp_Rencana, Status_FollowUp, Catatan_FollowUp, Tanggal_FollowUp_Aktual

**Sheet Item_Saran** (1 baris = 1 komponen yang disarankan):
ID_Item, ID_Kunjungan, Nama_Komponen, Qty, Keterangan_Teknisi, Nomor_Part, Estimasi_Harga, Ketersediaan_Part, Keterangan_Partman, Foto_URL, Diisi_Teknisi_At, Diisi_Partman_At
