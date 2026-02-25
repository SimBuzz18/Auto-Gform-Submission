# Auto Google Form Submission (AutoForm Pro)

Aplikasi otomatisasi pengisian Google Form berbasis Excel menggunakan Python, Selenium, dan CustomTkinter. Aplikasi ini dirancang untuk memudahkan input data massal ke Google Form dengan antarmuka yang modern dan fitur keamanan yang ketat.

## Fitur Unggulan

*   **Parallel Multi-Core Engine**: Mendukung pengisian massal secara bersamaan menggunakan pararelisasi proses (multi-processing), bukan sekadar threading.
*   **Mode Headless**: Pilihan untuk menjalankan browser tanpa muncul di layar untuk menghemat sumber daya.
*   **Antarmuka Modern (GUI)**: Menggunakan `customtkinter` dengan terminal log dinamis untuk tiap worker.
*   **Pencocokan Cerdas (Smart Matching)**: Judul kolom di Excel tidak harus persis 100% sama dengan pertanyaan di Google Form.
*   **Validasi Keamanan & Input**:
    *   **Worker Validator**: Mencegah input jumlah worker yang melebihi kapasitas core hardware.
    *   **Respondent Skipping**: Sistem lebih tangguh; jika data tidak cocok (mismatch), worker melompati responden tersebut dan lanjut ke data berikutnya (tidak berhenti total).
*   **Audit Fail-Safe (Conditional)**: Mengumpulkan data responden yang gagal secara otomatis dan menampilkannya hanya setelah seluruh proses selesai jika terdapat error. Juga mengekspor data ke file Excel terpisah.
*   **Tombol STOP Global**: Menghentikan seluruh worker dan menutup browser seketika jika terjadi kesalahan.

## Persyaratan (Requirements)

Pastikan Anda sudah menginstall Python. Kemudian install library yang dibutuhkan:

```bash
pip install pandas selenium webdriver-manager customtkinter openpyxl
```

## Cara Penggunaan

1.  **Siapkan File Excel**:
    *   Baris pertama harus berisi **Judul Pertanyaan** (usahakan mirip dengan di Google Form).
    *   Baris selanjutnya berisi data jawaban.
2.  **Jalankan Aplikasi**:
    Buka terminal/cmd di folder project, lalu jalankan:
    ```bash
    python ui.py
    ```
    *(Jangan jalankan `AutoForm.py` secara langsung, itu hanya file logika)*.
3.  **Di Dalam Aplikasi**:
    *   Paste **Link Google Form** Anda.
    *   Pilih **File Excel** Anda.
    *   Atur **Jumlah Worker** (Default: 75% Core PC).
    *   Pilih **Mode Headless** jika ingin bekerja di background.
    *   Klik tombol **START**.
    *   Pantau progress di terminal dinamis per worker.

## Struktur File

*   `ui.py`: Antarmuka GUI dan Orchestrator Pararel (Jalankan file ini).
*   `formLogic.py`: Mesin inti otomatisasi (Worker logic).
*   `CHANGELOG.md`: Catatan versi dan perubahan fitur.
*   `README.md`: Dokumentasi ini.

## Catatan Penting

*   **Format Checkbox**: Jika satu pertanyaan punya banyak jawaban (checkbox), pisahkan jawaban di Excel dengan koma (`,`). Contoh: `Nasi Goreng, Mie Goreng`.
*   **Jaringan**: Pastikan koneksi internet stabil karena Selenium membutuhkan akses real-time ke Google Form.

---
**Disclaimer**: Gunakan alat ini dengan bijak dan bertanggung jawab. Penulis tidak bertanggung jawab atas penyalahgunaan alat ini untuk spamming atau tindakan merugikan lainnya.
