# Auto Google Form Submission (AutoForm Pro)

Aplikasi otomatisasi pengisian Google Form berbasis Excel menggunakan Python, Selenium, dan CustomTkinter. Aplikasi ini dirancang untuk memudahkan input data massal ke Google Form dengan antarmuka yang modern dan fitur keamanan yang ketat.

## Fitur Unggulan

*   **Antarmuka Modern (GUI)**: Menggunakan `customtkinter` untuk tampilan yang bersih dan mudah digunakan (tidak perlu edit kodingan).
*   **Log Terminal Real-time**: Memantau proses pengisian secara langsung di jendela aplikasi.
*   **Pencocokan Cerdas (Smart Matching)**: Judul kolom di Excel tidak harus persis 100% sama dengan pertanyaan di Google Form (sistem pencocokan parsial).
*   **Validasi Ketat (Safety First)**:
    *   **Cek Pertanyaan**: Jika ada pertanyaan di Form yang tidak ditemukan di Excel, proses otomatis **BERHENTI**.
    *   **Cek Jawaban**: Jika jawaban di Excel (untuk Pilihan Ganda/Checkbox) tidak tersedia di opsi Form, proses otomatis **BERHENTI**.
*   **Tombol STOP/CANCEL**: Anda bisa memberhentikan proses kapan saja dengan aman.
*   **Browser Persisten**: Browser tidak akan tertutup otomatis setelah selesai, memungkinkan Anda mengecek hasil input.

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
    *   Klik tombol **START**.
    *   Duduk manis dan biarkan bot bekerja!

## Struktur File

*   `ui.py`: File utama aplikasi (Jalankan yang ini).
*   `AutoForm.py`: File logika/backend (Jantung aplikasi).
*   `README.md`: Dokumentasi ini.

## Catatan Penting

*   **Format Checkbox**: Jika satu pertanyaan punya banyak jawaban (checkbox), pisahkan jawaban di Excel dengan koma (`,`). Contoh: `Nasi Goreng, Mie Goreng`.
*   **Jaringan**: Pastikan koneksi internet stabil karena Selenium membutuhkan akses real-time ke Google Form.

---
**Disclaimer**: Gunakan alat ini dengan bijak dan bertanggung jawab. Penulis tidak bertanggung jawab atas penyalahgunaan alat ini untuk spamming atau tindakan merugikan lainnya.
