# =============================================================================
# ui.py — Entry Point Aplikasi AutoGForm Automation
# =============================================================================
# ARSITEKTUR SPAWN-SAFE:
#
# Pada Windows, multiprocessing menggunakan metode 'spawn'. Setiap kali
# multiprocessing.Process dibuat, Python MEMBUAT ULANG proses Python baru
# dan me-reimport file __main__ (yaitu file ini, ui.py) dari awal.
#
# MASALAH LAMA: import customtkinter/tkinter di top-level menyebabkan child
# process mencoba menginisialisasi Tk tanpa display → hang/crash diam-diam,
# sehingga worker tidak pernah menjalankan script.
#
# SOLUSI: ui.py HANYA berisi guard multiprocessing. Semua GUI (AutoFormApp)
# dipindahkan ke app_gui.py dan di-import HANYA di dalam blok __main__.
# Child process spawn me-reimport ui.py → langsung berhenti (tidak ada GUI).
# =============================================================================
import os
import sys
import multiprocessing

# BUGFIX: PyInstaller --noconsole multiprocessing
# multiprocessing.spawn akan mencoba menulis ke stdout/stderr yang bernilai None
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

# freeze_support() HARUS dipanggil di sini (sebelum logika apapun)
# agar PyInstaller tidak me-loop saat spawn di mode --onefile / --noconsole
multiprocessing.freeze_support()


if __name__ == "__main__":
    # Import GUI HANYA di sini — tidak akan dieksekusi oleh child process spawn
    from app_gui import AutoFormApp

    app = AutoFormApp()
    app.mainloop()
