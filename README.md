# Portal Filter Tenaga Ahli HRD

Aplikasi web berbasis **Streamlit** untuk membantu HRD melakukan pencarian, sortir, dan filter data tenaga ahli dari **Google Sheets**.

## Tujuan Aplikasi

Aplikasi ini membantu HRD mencari personil berdasarkan kriteria seperti:

- Kota/Kabupaten
- Provinsi
- Keahlian/SKK
- Status SKA
- Pendidikan/Ijazah
- Tahun lulus
- Penerbit SKA/SKK
- Nama, NIK, NPWP, atau keyword umum

Output dapat didownload sebagai:

- CSV
- PDF

---

## Arsitektur Sistem

```text
Google Sheets multi-sheet
        ↓
1 export XLSX link
        ↓
Streamlit Community Cloud
        ↓
Auto-cleaning data
        ↓
Filter / sort / download hasil
```

Aplikasi berjalan 100% di layanan gratis:

```text
GitHub                 = menyimpan source code
Streamlit Community Cloud = menjalankan aplikasi web
Google Sheets          = sumber data utama
```

---

## Struktur Repository

Struktur repository minimal:

```text
hr-data-filter/
├── app.py
├── requirements.txt
└── .streamlit/
    └── config.toml
```

File/folder yang **tidak boleh** di-commit ke GitHub:

```text
*.xlsx
*.xls
master_database.csv
uploaded_excels/
database_backups/
cleaning_reports/
dokumen_pelamar/
service-account.json
secrets.toml
```

---

## Google Sheets Source

Aplikasi membaca Google Sheet melalui link export XLSX.

Contoh URL Google Sheet:

```text
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit?usp=sharing
```

Link export XLSX yang dipakai aplikasi:

```text
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/export?format=xlsx
```

Contoh secrets:

```toml
[google_sheet]
xlsx_export_url = "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/export?format=xlsx"
refresh_seconds = 60
```

### Syarat Sharing Google Sheet

Karena aplikasi ini dibuat 100% free dan tidak memakai Google Cloud service account, Google Sheet harus bisa dibaca dari link:

```text
Share → Anyone with the link → Viewer
```

Jangan beri akses Editor.

Catatan keamanan:

> Siapa pun yang memiliki link Google Sheet dapat membaca data. Pastikan link hanya dibagikan ke pihak yang berwenang.

---

## Streamlit Cloud Setup

### 1. Push kode ke GitHub

```bash
git add app.py requirements.txt .streamlit/config.toml
git commit -m "Deploy HR data filter app"
git push
```

### 2. Deploy di Streamlit Community Cloud

Isi:

```text
Repository      : escdatahrd/hr-data-filter
Branch          : main
Main file path  : app.py
```

### 3. Tambahkan Secrets

Di Streamlit Cloud:

```text
Manage app → Settings → Secrets
```

Isi:

```toml
[google_sheet]
xlsx_export_url = "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/export?format=xlsx"
refresh_seconds = 60
```

### 4. Reboot App

Setelah secrets disimpan:

```text
Manage app → Reboot app
```

---

## requirements.txt

Contoh isi `requirements.txt`:

```txt
streamlit==1.37.1
pandas==2.2.2
numpy==1.26.4
openpyxl==3.1.5
XlsxWriter==3.2.0
fpdf2==2.7.9
requests==2.32.3
```

---

## Cara Kerja Aplikasi

### 1. Ambil Google Sheet

Aplikasi mengambil data dari URL export XLSX:

```text
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/export?format=xlsx
```

Lalu file tersebut dibaca sebagai workbook Excel multi-sheet.

### 2. Baca Multi-Sheet

Semua worksheet dibaca satu per satu.

Sheet yang diproses:

```text
ARS
SPL
MEKANIKAL
ELEK
MNJMN
NON TKNK
TAT ESC
TAT BPS
dan sheet personil lain
```

Sheet yang dilewati:

```text
data pendukung
KODE SKT
KODE SKA
Catatan
```

### 3. Auto-Cleaning

Engine cleaning melakukan:

- Standarisasi nama kolom.
- Deteksi header `NAMA`.
- Mapping kolom berbeda ke schema standar.
- Pemisahan `KOTA/KABUPATEN` dan `PROVINSI`.
- Deteksi dan pemindahan NIK.
- Deteksi dan pemindahan NPWP.
- Cleanup email dan hyperlink.
- Cleanup sumber data.
- Guard agar tanggal/tahun lulus tidak terbaca sebagai nomor telepon.
- Cleanup format tanggal.
- Deteksi status SKA aktif/expired.
- Retain hyperlink pada nama personil sebagai `LINK PERSONIL`.

### 4. Proses TAT ESC dan TAT BPS

Sheet `TAT ESC` dan `TAT BPS` tidak lagi dilewati.

Sistem mengekstrak data dari format TAT dan menyamakannya dengan schema utama.

Contoh:

```text
Nama                    → NAMA
IJASA DAN KELULUSAN     → JENIS IJAZAH
S1 Sipil 2017           → JENIS IJAZAH = S1 Sipil, TAHUN LULUS = 2017
SKA/SKK ... YANG DIMILIKI → KEAHLIAN
Exp : 10 April 2028     → TGL EXPIRED SKA
```

---

## Kolom Utama

Kolom yang ditampilkan dalam hasil filter:

```text
NO
NAMA
LINK PERSONIL
JENIS IJAZAH
KEAHLIAN
TGL EXPIRED SKA
STATUS SKA
KOTA/KABUPATEN
PROVINSI
TAHUN LULUS IJAZAH
NO NIK
NO NPWP
NO. TELP
EMAIL
SUMBER
KATEGORI_ASAL
```

---

## Logic Pendidikan / Ijazah

Aturan prioritas pendidikan:

```text
1. JENIS IJAZAH adalah sumber utama.
2. STRATA hanya dipakai jika JENIS IJAZAH kosong.
```

Contoh:

```text
JENIS IJAZAH = S1 Sipil
STRATA       = D3
```

Maka sistem memakai:

```text
S1 Sipil
```

Bukan `D3`.

Jika:

```text
JENIS IJAZAH = kosong
STRATA       = D3 Sipil
```

Maka sistem memakai:

```text
D3 Sipil
```

### Contoh Filter Pendidikan

Input:

```text
sipil d4 d3
```

Dimaknai sebagai:

```text
(Sipil + D4) ATAU (Sipil + D3)
```

Input:

```text
S1, D3
```

Dimaknai sebagai:

```text
S1 ATAU D3
```

Input:

```text
S1 serta D3
```

Dimaknai sebagai:

```text
S1 ATAU D3
```

---

## Logic Kota / Provinsi

Kolom lokasi dipisah:

```text
KOTA/KABUPATEN
PROVINSI
```

Tidak digabung lagi sebagai `DOMISILI` di hasil tabel.

Filter `Kota/Provinsi` tetap mencari ke dua kolom sekaligus.

Contoh input:

```text
Manado, Sulut, Bali
```

Dimaknai sebagai:

```text
Manado ATAU Sulut ATAU Bali
```

---

## Logic Keahlian / SKK

Filter keahlian mendukung banyak keyword.

Contoh:

```text
Jalan, Gedung, Arsitek
```

Dimaknai sebagai:

```text
Jalan ATAU Gedung ATAU Arsitek
```

Input seperti:

```text
SKK Jalan
```

tetap dicocokkan ke data seperti:

```text
Ahli Teknik Jalan - Madya Jenjang 8
Ahli Teknik Jalan - Utama Jenjang 9
```

---

## Link Personil ke Synology

Jika cell `NAMA` di Google Sheet memiliki hyperlink ke Synology, aplikasi akan mengambil hyperlink tersebut dari export XLSX.

Hasilnya tampil sebagai kolom:

```text
LINK PERSONIL
```

Di tabel, kolom tersebut muncul sebagai link:

```text
Buka
```

Catatan:

> Link hanya bisa terbaca jika hyperlink tersimpan sebagai hyperlink cell, bukan hanya teks URL biasa.

---

## Refresh Data

Aplikasi membaca data dari Google Sheet dengan cache.

Default:

```text
refresh_seconds = 60
```

Artinya data akan otomatis diperbarui maksimal sekitar 60 detik.

Untuk refresh langsung:

```text
Klik tombol Ambil Data
```

---

## Export Data

Aplikasi mendukung export:

```text
Download CSV
Download PDF
```

CSV cocok untuk olah data lanjutan.  
PDF cocok untuk laporan ringkas atau dibagikan ke user/client.

---

## Troubleshooting

### 1. App masih menampilkan versi lama

Penyebab:

- File `app.py` di GitHub belum terganti.
- Streamlit belum redeploy.
- Cache app masih lama.

Solusi:

```text
1. Pastikan app.py terbaru sudah di-push ke GitHub.
2. Streamlit Cloud → Manage app → Reboot app.
3. Klik Ambil Data.
```

---

### 2. Data tidak berubah setelah Google Sheet diedit

Penyebab:

- Cache masih aktif.
- Belum lewat 60 detik.
- Google Sheet belum tersimpan.
- Link export tidak bisa dibaca.

Solusi:

```text
Klik Ambil Data
```

Jika masih sama, cek:

```text
Share → Anyone with the link → Viewer
```

---

### 3. Error Google Sheet tidak bisa diakses

Cek secrets:

```toml
[google_sheet]
xlsx_export_url = "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/export?format=xlsx"
refresh_seconds = 60
```

Pastikan URL bukan link `/edit`, tapi `/export?format=xlsx`.

---

### 4. Hasil filter kosong

Cek apakah keyword sudah sesuai.

Contoh yang benar:

```text
Kota/Provinsi      : Manado, Sulut, Bali
Pendidikan/Ijazah  : sipil d4 d3
Keahlian/SKK       : Jalan, Gedung
Status SKA         : Aktif
```

Jika filter terlalu banyak diisi, sistem mencari data yang memenuhi semua filter sekaligus.

---

### 5. S1 masih ikut saat cari D3/D4

Pastikan sudah menggunakan versi terbaru yang menerapkan:

```text
Jenis Ijazah priority
```

Aturannya:

```text
JENIS IJAZAH diprioritaskan.
STRATA hanya dipakai jika JENIS IJAZAH kosong.
```

---

## Git Workflow

Setelah mengubah kode:

```bash
git add app.py requirements.txt .streamlit/config.toml
git commit -m "Update HR data filter app"
git push
```

Streamlit Cloud biasanya redeploy otomatis.

Jika tidak:

```text
Manage app → Reboot app
```

---

## Batasan Sistem

Aplikasi ini dibuat untuk mode gratis dan ringan.

Tidak termasuk:

```text
database permanen
login user
Google Cloud service account
backup file
upload Excel manual
dokumen pelamar
sinkronisasi Synology
```

Aplikasi ini fokus pada:

```text
Google Sheet live source
cleaning otomatis
filter/sort
export CSV/PDF
```

---

## Catatan Keamanan

Karena tidak memakai Google Cloud service account, Google Sheet harus dapat dibaca via link.

Konsekuensi:

```text
Anyone with the link → Viewer
```

Artinya siapa pun yang memiliki link Google Sheet dapat membaca data.

---

## Ringkasan

Aplikasi ini membaca data langsung dari Google Sheets, membersihkan data secara otomatis, lalu menyediakan filter/sort cepat untuk HRD. HRD cukup mengedit Google Sheet, kemudian aplikasi akan membaca pembaruan tersebut secara berkala. Tidak ada upload manual, tidak ada dokumen pelamar, dan tidak perlu server tambahan.
