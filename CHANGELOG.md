# Changelog

Semua perubahan signifikan pada proyek **AutoForm Pro** akan dicatat di file ini.

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
