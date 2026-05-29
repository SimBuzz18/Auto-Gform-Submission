# Auto Google Form Submission (AutoForm Pro)

> **Versi saat ini: v1.4.0**

Aplikasi otomatisasi pengisian Google Form berbasis Excel menggunakan Python, Selenium, dan CustomTkinter. Dirancang untuk input data massal ke Google Form secara paralel dengan antarmuka modern, mesin kill yang aman, dan dukungan sinkronisasi timestamp organik.

---

## Fitur Unggulan

### Mesin Otomatisasi
- **Parallel Multi-Core Engine**: Pengisian massal secara bersamaan menggunakan multiprocessing — bukan sekadar threading. Defaultnya menggunakan 75% core CPU.
- **7 Tipe Input Didukung** (dengan urutan deteksi otomatis):
  | Tipe | Format Excel |
  |------|-------------|
  | Kisi-kisi Pilihan Ganda (MCQ Grid) | `Baris1:Jawaban;Baris2:Jawaban` |
  | Petak Kotak Centang (Checkbox Grid) | `Baris1:Kol1\|Kol2;Baris2:Kol3` |
  | Radio Button / Skala Linear / Rating | Nilai teks atau angka |
  | Dropdown | Nilai teks |
  | Tanggal | `DD/MM/YYYY` |
  | Waktu | `HH:MM` |
  | Text / Textarea | Teks bebas |
- **Smart Matching**: Nama kolom di Excel dicocokkan dengan pertanyaan Google Form secara case-insensitive dan whitespace-tolerant.
- **Respondent Skip & Network Retry**: Membedakan data mismatch (langsung skip) vs. error jaringan (retry hingga 3x).

### Keamanan & Stabilitas
- **Process Tree Kill**: Tombol STOP mematikan seluruh process tree secara paksa — Python worker, `chromedriver`, dan browser Chrome — menggunakan `psutil`. Browser tidak akan tertinggal sebagai orphan process.
- **PID Registry**: Worker mendaftarkan PID browser-nya ke GUI segera setelah browser terbuka, memungkinkan kill yang presisi.
- **Spawn-Safe Architecture**: Entry point `ui.py` menggunakan guard `if __name__ == "__main__"` dan `multiprocessing.freeze_support()` untuk mencegah crash saat multiprocessing spawn di Windows.

### Sinkronisasi Timestamp *(Opsional)*
- Patch otomatis kolom `Timestamp` di Google Sheets setelah setiap respons dikirim, menggunakan nilai timestamp dari Excel organik.
- Jika kolom timestamp tidak ada atau kosong, fallback ke waktu submit aktual.
- Autentikasi via Google Service Account JSON (tidak membutuhkan login manual).

### Antarmuka (GUI)
- Mode Headless untuk pengisian di background.
- Terminal log dinamis per worker.
- Panel audit responden gagal (ekspor otomatis ke `Responden Gagal.xlsx`).

---

## Persyaratan (Requirements)

Python 3.8+ dan semua dependensi berikut:

```bash
pip install -r requirements.txt
```

Atau install manual:
```bash
pip install pandas selenium webdriver-manager customtkinter openpyxl psutil gspread google-auth
```

---

## Cara Penggunaan

### 1. Siapkan File Excel

Baris pertama = **Judul Pertanyaan** (usahakan mirip dengan teks pertanyaan di Google Form).

**Contoh format kolom khusus:**

| Timestamp | Nama Lengkap | Kualitas Produk | Fitur yang Dipilih | Tanggal Lahir | Jam Masuk |
|-----------|-------------|----------------|-------------------|---------------|-----------|
| 29/05/2024 09:30:00 | Budi | Sangat Baik | Fitur A,Fitur B | 15/03/2000 | 08:30 |

- **MCQ Grid**: `"Ketepatan:Setuju;Kecepatan:Sangat Setuju"`
- **Checkbox Grid**: `"Desain:Simpel|Modern;Fitur:Lengkap"`
- **Checkbox biasa**: `"Opsi A,Opsi B"` (pisahkan dengan koma)

### 2. Jalankan Aplikasi

```bash
python ui.py
```

> Jangan jalankan `app_gui.py` atau `formLogic.py` langsung — keduanya adalah modul pendukung.

### 3. Konfigurasi di Aplikasi

1. Paste **Link Google Form**
2. Pilih **File Excel**
3. Atur **Jumlah Worker** (default: 75% core CPU)
4. *(Opsional)* Aktifkan **Mode Headless**
5. *(Opsional)* Aktifkan **Sinkronisasi Timestamp** — isi Spreadsheet ID dan pilih Service Account JSON
6. Klik **START** dan pantau terminal per worker

---

## Setup Sinkronisasi Timestamp

Fitur ini membutuhkan Google Service Account. Langkah setup:

### A. Buat Service Account
1. Buka [Google Cloud Console](https://console.cloud.google.com/)
2. `APIs & Services` → `Credentials` → `Create Credentials` → `Service Account`
3. Aktifkan **Google Sheets API** dan **Google Drive API**
4. Download key dalam format JSON

### B. Bagikan Spreadsheet ke Service Account
1. Buka spreadsheet respons Google Form
2. Klik `Share` → masukkan email service account (format: `xxx@project.iam.gserviceaccount.com`)
3. Beri role **Editor**

### C. Ambil Spreadsheet ID
Dari URL spreadsheet:
```
https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit
```

---

## Struktur File

```
.
├── ui.py           # Entry point (jalankan file ini)
├── app_gui.py      # Definisi GUI & Orchestrator
├── formLogic.py    # Mesin inti otomatisasi (Worker logic)
├── requirements.txt
├── CHANGELOG.md
└── README.md
```

---

## Catatan Penting

- **Format Checkbox**: Pisahkan pilihan dengan koma. Contoh: `Nasi Goreng,Mie Goreng`
- **Jaringan**: Pastikan koneksi internet stabil — Selenium membutuhkan akses real-time ke Google Form
- **Timestamp Sync**: Karena Google Form mencatat timestamp di server-side, fitur ini bekerja dengan me-patch cell Timestamp di Google Sheets **setelah** respons masuk. Diperlukan `google-auth` dan `gspread`
- **Race Condition Multi-Worker**: Dengan lebih dari 1 worker, ada kemungkinan kecil konflik saat patching timestamp jika dua worker mengirim form secara bersamaan dalam <1 detik. Risiko ini diminimalkan oleh jeda 0.5 detik antar spawn worker

---

**Disclaimer**: Gunakan alat ini dengan bijak dan bertanggung jawab. Penulis tidak bertanggung jawab atas penyalahgunaan alat ini untuk spamming atau tindakan merugikan lainnya.
