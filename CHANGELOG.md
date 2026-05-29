# Changelog

Semua perubahan signifikan pada proyek **AutoForm Pro** akan dicatat di file ini.

## [1.4.1] - 2026-05-29

### Fixed

- **Page Load Sync**: Menambahkan mekanisme penantian dinamis `document.readyState == 'complete'` dan deteksi *staleness* pada transisi halaman untuk memastikan browser 100% terload sebelum worker mengisi form.
- **Dynamic Stabilization Loop**: Menambahkan pemantauan dinamis terhadap jumlah elemen pertanyaan pada halaman untuk mencegah data ter-skip akibat render JavaScript/React yang lambat.
- **Form Submission Verification**: Menambahkan verifikasi halaman konfirmasi pasca-pengiriman (menunggu hingga 15 detik untuk memvalidasi teks sukses dan menghilangnya pertanyaan dari DOM) guna mencegah penutupan browser prematur yang menyebabkan submit palsu.
- **Hybrid Click Helper (`_click_element`)**: Menerapkan metode klik hibrida (native click dengan fallback ke JavaScript click setelah memposisikan elemen di tengah layar) di seluruh komponen form untuk menjamin event click terdaftar di React handler.
- **Explicit Headless Resolution**: Menambahkan konfigurasi `--window-size=1920,1080` untuk memastikan tata letak halaman termuat utuh dalam mode headless pararel.
- **Removed Topmost Window Lock**: Menghapus `app.attributes('-topmost', True)` pada `ui.py` dan `app_gui.py` agar aplikasi form automation tidak menutupi jendela aplikasi aktif lainnya di Windows taskbar.
- **Selenium Manager Integration**: Menggantikan `webdriver_manager` (yang sering kali hang di latar belakang dengan status 'Menyiapkan...') dengan Selenium Manager bawaan Selenium 4.6+, sehingga tombol START langsung aktif seketika saat aplikasi dibuka.

## [1.4.0] - 2026-05-29

### Added

- **Process Tree Kill (Stop Button)**: Menekan tombol "STOP" sekarang mematikan **seluruh process tree** secara paksa — termasuk `chromedriver` dan browser Chrome yang di-spawn oleh Selenium — menggunakan `psutil`. Sebelumnya hanya Python worker yang di-terminate, menyebabkan browser Chrome tetap berjalan sebagai orphan process.
- **PID Registry via Queue**: Worker mendaftarkan PID `chromedriver`-nya ke GUI melalui sinyal internal `[PID_REGISTER]` segera setelah browser berhasil dibuka. GUI menyimpan PID ini untuk digunakan saat kill diperlukan.
- **Dropdown Handler**: Pendeteksian dan pengisian dropdown Google Form (custom `div[role='listbox']`). Mekanisme: klik trigger → tunggu popup muncul via `WebDriverWait` → klik opsi yang cocok → fallback Escape jika opsi tidak ditemukan.
- **Multiple Choice Grid Handler (Kisi-kisi Pilihan Ganda)**: Deteksi otomatis berdasarkan keberadaan `> 1` elemen `div[role='radiogroup']` di dalam satu pertanyaan. Format Excel: `NamaBaris1:JawabanKolom;NamaBaris2:JawabanKolom`.
- **Checkbox Grid Handler (Petak Kotak Centang)**: Deteksi via heuristik sampling `aria-label` (format `"NamaBaris, NamaKolom"`). Format Excel: `Baris1:Kolom1|Kolom2;Baris2:Kolom3`.
- **Date Handler (Tanggal)**: Pengisian otomatis input tanggal Google Form (3 field terpisah: Day/Month/Year). Mendukung format Excel: `DD/MM/YYYY`, `YYYY-MM-DD`, `DD-MM-YYYY`, dan objek `datetime` pandas.
- **Time Handler (Waktu)**: Pengisian otomatis input waktu (field Hour/Minute terpisah). Format Excel: `HH:MM`.
- **Linear Scale & Rating**: Ditangani oleh Radio Button handler (Section C) — tipe ini menggunakan elemen `div[role='radio']` dengan `data-value` numerik. Tidak memerlukan handler terpisah.
- **Timestamp Synchronization (Google Sheets API)**: Fitur opsional untuk menyinkronkan timestamp respons di Google Sheets dengan timestamp dari data Excel organik. Mekanisme: setelah form berhasil disubmit, Sheets API memperbarui cell kolom A (Timestamp) di baris terakhir menggunakan nilai dari kolom timestamp Excel. Jika kolom timestamp tidak ada atau kosong, menggunakan waktu sekarang sebagai fallback.
- **Service Account Integration**: Autentikasi ke Google Sheets menggunakan Service Account JSON (`google-auth`). Koneksi `gspread` dibuat sekali per worker (lazy-init, reused).
- **Timestamp Column Auto-detection**: Deteksi otomatis nama kolom timestamp di Excel (alias yang dikenali: `Timestamp`, `Stempel waktu`, `Waktu`, `Waktu Submit`, `Tanggal Submit`, `Tanggal`, `Time`, `Submit Time`, `Submit At`).
- **Timestamp Sync UI Section**: Bagian baru di panel kiri GUI — checkbox aktifkan fitur, input Spreadsheet ID, dan file browser untuk Service Account JSON.

### Changed

- **Input Handler Order Refactor**: Urutan deteksi tipe input diubah secara fundamental. Grid handler (MCQ Grid & Checkbox Grid) **wajib** diperiksa sebelum Radio/Checkbox biasa karena keduanya mengandung elemen yang sama (`div[role='radio']` / `div[role='checkbox']`). Urutan baru: `A. MCQ Grid → B. Checkbox Grid → C. Radio → D. Dropdown → E. Date → F. Time → G. Text`.
- **`worker_launcher` Signature**: Menambahkan parameter `spreadsheet_id` dan `creds_path` (keduanya opsional, default `None`).
- **`logic.__init__` Signature**: Menambahkan parameter `spreadsheet_id`, `creds_path`, dan instance var `_gsheet` (lazy singleton).

### Fixed

- **Stale Element in Dropdown**: Menambahkan `StaleElementReferenceException` handler saat iterasi opsi dropdown yang bisa berubah saat popup terbuka.
- **Log Listener Guard**: Menambahkan `except Exception: pass` sebagai guard umum di `log_listener` agar thread tidak mati diam-diam akibat error tak terduga.

### Dependencies Added

- `psutil` — process tree kill
- `gspread` — Google Sheets API client
- `google-auth` — autentikasi Service Account Google

---

## [1.2.9] - 2026-02-25

### Added
- **Force Stop Feature**: Menekan tombol "STOP" sekarang akan langsung mematikan (`terminate`) seluruh proses worker secara instan.

### Fixed
- **Threading RuntimeError**: Memperbaiki `RuntimeError: main thread is not in main loop` yang muncul saat Orchestrator mencoba memperbarui UI secara langsung tanpa melalui antrian `self.after()`.

## [1.2.8] - 2026-02-25

### Fixed
- **Worker Startup Stability**: Menangani masalah browser "stuck" dengan cara men-download driver Chrome satu kali di main process dan memberikan jeda 0.5 detik antar peluncuran browser.

## [1.2.7] - 2026-02-25

### Added
- **Audit Data Sorting**: Data pada panel audit dan file rekap `Responden Gagal.xlsx` kini otomatis diurutkan berdasarkan nomor baris (terkecil ke terbesar) untuk memudahkan pengecekan manual.

## [1.2.6] - 2026-02-25

### Fixed
- **Audit Panel KeyError**: Memperbaiki `KeyError: 'Baris'` di `show_audit_panel`. Sekarang data audit dipastikan memiliki kunci standar (`Baris`, `Nama`, `Alasan`) meskipun datanya berasal dari kolom Excel yang berbeda-beda.

## [1.2.5] - 2026-02-25

### Fixed
- **Synchronous Audit Collection**: Memperbaiki secara total masalah race condition pada panel audit. Sekarang data error diproses langsung di thread listener (bukan lagi menunggu callback UI), sehingga data dijamin sudah siap sebelum ringkasan audit ditampilkan atau diekspor.

## [1.2.4] - 2026-02-25

### Fixed
- **Missing Import**: Memperbaiki `NameError` di `ui.py` karena modul `time` belum di-import namun digunakan di fungsi `run_orchestrator`.

## [1.2.3] - 2026-02-25

### Fixed
- **Audit Panel Race Condition**: Memperbaiki masalah di mana panel audit terkadang tidak muncul karena script selesai lebih cepat daripada proses penulisan log di UI. Sekarang Orchestrator menunggu listener log hingga benar-benar kosong sebelum menampilkan rekap.
- **Full-Row Audit Robustness**: Memastikan data baris lengkap tercatat dengan benar meskipun terdapat variasi index.

## [1.2.2] - 2026-02-25

### Added
- **Full-Row Audit System**: File rekap `Responden Gagal.xlsx` kini berisi **seluruh kolom asli** dari file Excel yang diimport (copas baris lengkap) untuk memudahkan analisis manual.
- **Enhanced Exception Propagation**: Error interaksi (seperti elemen tidak bisa diklik) sekarang otomatis memicu Skip, mencegah penekanan tombol 'Kirim' pada form yang belum terisi lengkap.
- **Queue Draining Logic**: Listener log sekarang memastikan seluruh pesan dari worker diproses sebelum menampilkan panel audit.

### Fixed
- **Network Resilience Handling**: Mempertahankan mekanisme Retry untuk error jaringan/selenium, namun tetap mencatat data baris lengkap ke audit jika seluruh percobaan gagal.

## [1.2.1] - 2026-02-25

### Fixed
- **Skip Respondent Exception Swallowing**: Memperbaiki bug di mana worker tetap melanjutkan pengisian (dan menekan Kirim) meskipun data tidak ditemukan/tidak cocok. Sekarang worker akan langsung berhenti memproses responden tersebut saat error ditemukan.

## [1.2.0] - 2026-02-25

### Added
- **Conditional Audit Panel**: Panel audit (Ringkasan Gagal) hanya muncul setelah seluruh proses selesai dan jika terdapat data error.
- **Respondent Skipping Logic**: Worker sekarang tidak terhenti jika data tidak cocok (missing question/option); responden dilewati, dilaporkan ke audit, dan proses lanjut ke data berikutnya.

### Fixed
- **UI Layout Handling**: Reset state UI setiap kali tombol START ditekan untuk mencegah tumpang tindih layout terminal.

## [1.1.0] - 2026-02-25

### Added
- **Multi-Process Parallel Engine**: Dukungan pengisian form secara pararel menggunakan hingga 75% core CPU.
- **Dynamic Worker Configuration**: User bisa mengatur jumlah worker (browser) secara manual melalui UI.
- **Headless Mode Toggle**: Opsi untuk menjalankan browser di background atau terlihat.
- **Real-time Worker Terminals**: Tampilan log terpisah untuk tiap worker secara dinamis.
- **Error Summary Panel**: Tabel rekap responden yang gagal secara real-time.
- **Automated Audit Export**: Ekspor otomatis ke `Responden Gagal.xlsx` jika ada kegagalan pengisian.
- **Strict Input Validation**: Validasi angka dan range pada input jumlah worker.

### Optimized
- **WebDriverWait Implementation**: Migrasi dari jeda statis (`time.sleep`) ke dinamis (`WebDriverWait`) untuk kecepatan maksimal.
- **Smart Data Chunking**: Pembagian beban kerja yang merata antar worker.
- **Resource Management**: Pembersihan driver Chrome yang lebih bersih untuk mencegah memory leak.

## [1.0.0] - 2026-02-24

### Added
- Inisialisasi proyek AutoForm Pro.
- GUI menggunakan CustomTkinter.
- Fitur Smart Matching kolom Excel ke Pertanyaan GForm.
- Dukungan Radio Button, Checkbox, dan Text Input.
