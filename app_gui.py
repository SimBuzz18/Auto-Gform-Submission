# =============================================================================
# app_gui.py — Definisi AutoFormApp GUI
# =============================================================================
# File ini HANYA di-import oleh ui.py di dalam blok 'if __name__ == "__main__"'.
# Memisahkan definisi GUI dari entry point adalah KUNCI agar child process
# multiprocessing (spawn) tidak me-reimport customtkinter/tkinter dan crash.
# =============================================================================
import os
import sys
import time
import threading
import math
import re
import queue
import signal
import multiprocessing
from datetime import datetime

try:
    import psutil
    _PSUTIL_AVAILABLE = True
except ImportError:
    _PSUTIL_AVAILABLE = False

import pandas as pd
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk
from webdriver_manager.chrome import ChromeDriverManager

from formLogic import logic, worker_launcher


# --- GUI APP ---
class AutoFormApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Window Config
        self.title("AutoGForm Automation")
        self.geometry("1200x800")
        # Gunakan .after() untuk mengeksekusi 'zoomed' setelah UI selesai dirender
        self.after(0, lambda: self.state('zoomed'))
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Layout Config
        self.grid_columnconfigure(1, weight=1) # Kanan expand
        self.grid_rowconfigure(0, weight=1)    # Full height

        # State Variables
        self.is_running = False
        self.processes = []
        self.worker_pids = {}  # {worker_id: [pid_chrome, pid_chromedriver, ...]}
        self.queue_log = multiprocessing.Queue()
        self.stop_event = multiprocessing.Event()
        self.link_var = tk.StringVar()
        self.file_path_var = tk.StringVar()
        self.error_data = [] # Koleksi data responden gagal: {'Baris': x, 'Nama': y, 'Alasan': z}
        self.cached_driver_path = None # Variabel konfigurasi untuk menyimpan path driver

        # New Settings Variables
        self.headless_var = tk.BooleanVar(value=True)
        self.total_cores = multiprocessing.cpu_count()
        self.default_workers = max(1, int(self.total_cores * 0.75))
        self.worker_var = tk.StringVar(value=str(self.default_workers))
        self.worker_var.trace_add("write", self._validate_worker_range)

        # Timestamp Sync Variables
        self.use_ts_sync_var    = tk.BooleanVar(value=False)
        self.spreadsheet_id_var = tk.StringVar()
        self.creds_path_var     = tk.StringVar()
        
        # --- LEFT PANEL (INPUTS) ---
        self.frame_left = ctk.CTkFrame(self, width=100, corner_radius=0)
        self.frame_left.grid(row=0, column=0, sticky="nswe", padx=0, pady=0)
        self.frame_left.grid_rowconfigure(10, weight=1) # Spacer bawah

        self._init_left_panel()

        # --- RIGHT PANEL (TERMINAL) ---
        self.frame_right = ctk.CTkFrame(self, corner_radius=0)
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=10, pady=10)
        self.frame_right.grid_rowconfigure(1, weight=1)
        self.frame_right.grid_columnconfigure(0, weight=1)

        self._init_right_panel()

        # Flag untuk menandai apakah Tk mainloop masih aktif.
        # Digunakan oleh _safe_after() untuk mencegah RuntimeError
        # saat thread memanggil self.after() setelah window di-destroy.
        self._destroyed = False
        # Override protocol WM_DELETE_WINDOW agar cleanup berjalan saat user menutup window
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # Mulai validasi ChromeDriver di background saat aplikasi dibuka
        self._safe_after(100, self._init_driver_background)

    def _init_driver_background(self):
        """Menyiapkan ChromeDriver di latar belakang secara otomatis via Selenium Manager."""
        def check_driver():
            try:
                self._safe_after(0, lambda: self.log_gui("[System] Menggunakan Selenium Manager bawaan (Chrome 115+ / Selenium 4.6+)..."))
                self.cached_driver_path = None
                self._safe_after(0, lambda: self.log_gui("[System] ChromeDriver siap (dikelola otomatis oleh Selenium Manager)\n"))
                self._safe_after(0, lambda: self.btn_start.configure(state="normal", text="START"))
            except Exception as e:
                self._safe_after(0, lambda: self.log_gui(f"[System] [!] Gagal: {e}\n"))
                self._safe_after(0, lambda: self.btn_start.configure(state="normal", text="START"))
        
        t = threading.Thread(target=check_driver, daemon=True)
        t.start()

    def _init_left_panel(self):
        lbl_title = ctk.CTkLabel(self.frame_left, text="AutoGForm Automation", font=ctk.CTkFont(size=20, weight="bold"))
        lbl_title.grid(row=0, column=0, padx=20, pady=(20, 10))

        # 1. Link Input
        lbl_link = ctk.CTkLabel(self.frame_left, text="Link Google Form:", anchor="w")
        lbl_link.grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")
        entry_link = ctk.CTkEntry(self.frame_left, textvariable=self.link_var, placeholder_text="Paste Link...")
        entry_link.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")

        # 2. File Input
        lbl_file = ctk.CTkLabel(self.frame_left, text="File Excel:", anchor="w")
        lbl_file.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="w")
        
        btn_browse = ctk.CTkButton(self.frame_left, text="Pilih File Excel", command=self.browse_file, fg_color="#333", hover_color="#555")
        btn_browse.grid(row=4, column=0, padx=20, pady=(0, 5), sticky="ew")
        
        lbl_path = ctk.CTkLabel(self.frame_left, textvariable=self.file_path_var, font=ctk.CTkFont(size=10), text_color="gray", wraplength=250)
        lbl_path.grid(row=5, column=0, padx=20, pady=(0, 10), sticky="w")

        # 3. Settings Panel
        self.frame_settings = ctk.CTkFrame(self.frame_left, fg_color="transparent")
        self.frame_settings.grid(row=6, column=0, padx=20, pady=10, sticky="ew")
        
        self.check_headless = ctk.CTkCheckBox(self.frame_settings, text="Mode Headless", variable=self.headless_var)
        self.check_headless.grid(row=0, column=0, sticky="w", pady=(0, 10))
        
        self.frame_worker = ctk.CTkFrame(self.frame_settings, fg_color="transparent")
        self.frame_worker.grid(row=1, column=0, sticky="w")
        
        lbl_worker = ctk.CTkLabel(self.frame_worker, text=f"Worker (Max {self.total_cores}):")
        lbl_worker.grid(row=0, column=0, sticky="w")
        
        # Registrasi validator untuk mencegah input non-angka
        vcmd = (self.register(self._only_digits), '%P')
        self.entry_worker = ctk.CTkEntry(self.frame_worker, textvariable=self.worker_var, width=50, validate="key", validatecommand=vcmd)
        self.entry_worker.grid(row=0, column=1, padx=5)

        # Control Buttons
        # Tombol dimatikan saat awal sampai ChromeDriver selesai divalidasi
        self.btn_start = ctk.CTkButton(self.frame_left, text="Menyiapkan...", command=self.start_automation, height=40, font=ctk.CTkFont(weight="bold"), state="disabled")
        self.btn_start.grid(row=7, column=0, padx=20, pady=(20, 10), sticky="ew")

        self.btn_stop = ctk.CTkButton(self.frame_left, text="STOP", command=self.stop_automation, height=40, fg_color="darkred", hover_color="#aa0000", state="disabled")
        self.btn_stop.grid(row=8, column=0, padx=20, pady=10, sticky="ew")

        # 4. Timestamp Sync Section
        sep = ctk.CTkFrame(self.frame_left, height=1, fg_color="#444")
        sep.grid(row=9, column=0, padx=20, pady=(5, 0), sticky="ew")

        self.frame_ts = ctk.CTkFrame(self.frame_left, fg_color="transparent")
        self.frame_ts.grid(row=10, column=0, padx=20, pady=(5, 10), sticky="ew")
        self.frame_ts.grid_columnconfigure(0, weight=1)

        # Judul section
        lbl_ts_title = ctk.CTkLabel(
            self.frame_ts, text="Sinkronisasi Timestamp",
            font=ctk.CTkFont(size=11, weight="bold"), anchor="w"
        )
        lbl_ts_title.grid(row=0, column=0, sticky="w", pady=(0, 4))

        # Checkbox aktifkan
        self.check_ts = ctk.CTkCheckBox(
            self.frame_ts, text="Aktifkan (butuh Service Account)",
            variable=self.use_ts_sync_var, font=ctk.CTkFont(size=11),
            command=self._toggle_ts_fields
        )
        self.check_ts.grid(row=1, column=0, sticky="w", pady=(0, 6))

        # Spreadsheet ID
        lbl_ss_id = ctk.CTkLabel(self.frame_ts, text="Google Spreadsheet ID:", anchor="w", font=ctk.CTkFont(size=11))
        lbl_ss_id.grid(row=2, column=0, sticky="w")
        self.entry_ss_id = ctk.CTkEntry(
            self.frame_ts, textvariable=self.spreadsheet_id_var,
            placeholder_text="Paste Spreadsheet ID...", state="disabled"
        )
        self.entry_ss_id.grid(row=3, column=0, sticky="ew", pady=(0, 6))

        # Service Account JSON
        lbl_creds = ctk.CTkLabel(self.frame_ts, text="Service Account JSON:", anchor="w", font=ctk.CTkFont(size=11))
        lbl_creds.grid(row=4, column=0, sticky="w")
        self.btn_creds = ctk.CTkButton(
            self.frame_ts, text="Pilih File JSON...",
            command=self.browse_creds, fg_color="#333", hover_color="#555",
            state="disabled", font=ctk.CTkFont(size=11)
        )
        self.btn_creds.grid(row=5, column=0, sticky="ew", pady=(0, 2))
        lbl_creds_path = ctk.CTkLabel(
            self.frame_ts, textvariable=self.creds_path_var,
            font=ctk.CTkFont(size=9), text_color="gray", wraplength=250, anchor="w"
        )
        lbl_creds_path.grid(row=6, column=0, sticky="w")

    def _init_right_panel(self):
        lbl_term = ctk.CTkLabel(self.frame_right, text="Terminal Output & Log", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_term.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        # Textbox (Terminal Utama / Orchestrator)
        self.main_terminal = ctk.CTkTextbox(self.frame_right, width=400, font=("Consolas", 12))
        self.main_terminal.grid(row=1, column=0, padx=10, pady=10, sticky="nswe")
        self.main_terminal.insert("0.0", "Siap menjalankan otomatisasi...\n")
        self.main_terminal.configure(state="disabled") 
        self.worker_terminals = {} # Tempat menyimpan kotak terminal per worker
        
        # NOTE: Panel Audit (Ringkasan Gagal) sekarang dibuat dinamis di show_audit_panel()
        # agar hanya muncul setelah proses selesai dan jika ada error.
        self.error_terminal = None 
    
    def setup_worker_terminals(self, worker_ranges):
        """Membuat kolom textboxes secara dinamis berdasarkan jumlah worker dan rentang data."""
        n_workers = len(worker_ranges)
        
        # Simpan log lama dari orchestrator
        old_logs = self.main_terminal.get("0.0", "end")
        
        # Bersihkan frame kanan
        for widget in self.frame_right.winfo_children():
            widget.destroy()

        # Atur ulang layout grid
        self.frame_right.grid_rowconfigure(1, weight=0) # Baris log utama (fixed)
        self.frame_right.grid_rowconfigure(3, weight=1) # Baris log worker (expand)

        # 1. Buat Terminal Orchestrator (Di atas)
        lbl_main = ctk.CTkLabel(self.frame_right, text="Orchestrator Log (Main)", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_main.grid(row=0, column=0, columnspan=max(1, n_workers), sticky="w", padx=10, pady=(10,0))
        
        self.main_terminal = ctk.CTkTextbox(self.frame_right, height=100, font=("Consolas", 10))
        self.main_terminal.grid(row=1, column=0, columnspan=max(1, n_workers), sticky="ew", padx=10, pady=(0, 10))
        self.main_terminal.configure(state="normal")
        self.main_terminal.insert("0.0", old_logs.strip() + "\n")
        self.main_terminal.see("end")
        self.main_terminal.configure(state="disabled")
        

        self.terminal_font_size = 6
        # 2. Buat Terminal Worker (Berjejer ke Samping)
        self.worker_terminals = {}
        for i, (start_n, end_n) in enumerate(worker_ranges):
            worker_id = i + 1
            self.frame_right.grid_columnconfigure(i, weight=1, uniform="col") 
            
            # Format Judul: Worker 1 - N: 1 - 25
            header_text = f"Worker {worker_id} - N: {start_n} - {end_n}"
            
            lbl_worker = ctk.CTkLabel(self.frame_right, text=header_text, font=ctk.CTkFont(size=12, weight="bold"))
            lbl_worker.grid(row=2, column=i, pady=(0,5))
            
            txt = ctk.CTkTextbox(self.frame_right, font=("Consolas", self.terminal_font_size), border_width=1, border_color="#555")
            txt.grid(row=3, column=i, sticky="nswe", padx=5, pady=(0, 10))
            txt.configure(state="disabled")
            self.worker_terminals[worker_id] = txt

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename:
            self.file_path_var.set(filename)
            self.log_gui(f"File dipilih: {os.path.basename(filename)}")

    def browse_creds(self):
        """File browser untuk memilih Service Account JSON."""
        filename = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if filename:
            self.creds_path_var.set(filename)
            self.log_gui(f"[TS] Service Account: {os.path.basename(filename)}")

    def _toggle_ts_fields(self):
        """Enable/disable field timestamp sync berdasarkan state checkbox."""
        state = "normal" if self.use_ts_sync_var.get() else "disabled"
        self.entry_ss_id.configure(state=state)
        self.btn_creds.configure(state=state)

    def _only_digits(self, P):
        """Mencegah karakter selain angka (digit)."""
        if P == "" or P.isdigit():
            return True
        return False

    def _validate_worker_range(self, *args):
        """Validasi range worker: minimal 1, maksimal total_cores."""
        val = self.worker_var.get()
        if val.isdigit():
            n = int(val)
            if n > self.total_cores:
                self.worker_var.set(str(self.total_cores))
            elif n < 1 and val != "": # Jangan paksa '1' kalau sedang dihapus total
                pass 
        elif val == "":
            pass # Biarkan kosong saat user hapus total

    def log_gui(self, message):
        """Append text to the GUI terminal in a thread-safe way."""
        # Update terminal error jika sudah dipicu tampilannya
        # message format: [ERROR_UI]|Baris|Nama|Alasan
        if message.startswith("[ERROR_UI]|"):
            parts = message.split("|")
            if len(parts) >= 4 and self.error_terminal:
                row_idx, nama, alasan = parts[1], parts[2], parts[3]
                self.error_terminal.configure(state="normal")
                self.error_terminal.insert("end", f"{row_idx.ljust(6)} | {nama.ljust(20)} | {alasan}\n")
                self.error_terminal.see("end")
                self.error_terminal.configure(state="disabled")
            return

        # B. Cek apakah teks mengandung identitas worker (contoh: "[Worker-1] Mengisi data...")
        match = re.match(r"^\[Worker-(\d+)\]\s*(.*)", message)
        
        if match:
            worker_id = int(match.group(1))
            text = match.group(2)
            # Arahkan ke kolom worker jika sudah dibuat, jika belum ke terminal utama
            if hasattr(self, 'worker_terminals') and worker_id in self.worker_terminals:
                target_term = self.worker_terminals[worker_id]
                msg_to_print = text
            else:
                target_term = self.main_terminal
                msg_to_print = message
        else:
            target_term = self.main_terminal
            msg_to_print = message

        if target_term:
            target_term.configure(state="normal")
            target_term.insert("end", f"{msg_to_print}\n")
            target_term.see("end")
            target_term.configure(state="disabled")

    def _safe_after(self, ms, func, *args):
        """Thread-safe wrapper untuk self.after().
        
        Mencegah RuntimeError 'main thread is not in main loop' yang terjadi
        ketika user menutup window saat thread orchestrator/listener masih aktif.
        """
        if self._destroyed:
            return
        try:
            self.after(ms, func, *args)
        except RuntimeError:
            # Window sudah di-destroy, abaikan
            pass

    def _on_close(self):
        """Handler saat user menutup window (klik tombol X).
        
        Menghentikan semua worker process secara paksa dan menandai
        bahwa Tk sudah tidak aktif agar _safe_after() tidak crash.
        """
        self._destroyed = True
        # Set stop event agar semua worker berhenti
        self.stop_event.set()
        # Terminate paksa semua proses worker yang masih hidup
        for p in self.processes:
            if p.is_alive():
                try:
                    p.terminate()
                except Exception:
                    pass
        self.is_running = False
        self.destroy()

    def start_automation(self):
        if self.is_running: return
        
        # Validasi akhir untuk input worker (jika kosong atau 0)
        try:
            val = self.worker_var.get().strip()
            if val == "" or int(val) < 1:
                self.worker_var.set("1")
        except:
            self.worker_var.set(str(self.default_workers))

        # Reset State & UI
        self.error_terminal = None 
        self.error_data = [] 
        self.current_df = None # Tempat simpan data asli untuk lookup baris

        link = self.link_var.get().strip()
        file = self.file_path_var.get().strip()
        
        if not link or not file:
            tk.messagebox.showwarning("Warning", "Link dan File Excel harus diisi!")
            return

        self.is_running = True
        # FIX: Hapus dead variable 'stop_requested' — sinyal stop yang aktif adalah stop_event
        # Bersihkan state dari run sebelumnya sebelum spawn worker baru
        self.stop_event.clear()
        self.worker_pids = {}  # Reset PID registry untuk run baru
        # Drain sisa pesan di queue dari run sebelumnya agar tidak terbaca listener baru
        while not self.queue_log.empty():
            try:
                self.queue_log.get_nowait()
            except queue.Empty:
                break
        # Simpan objek waktu untuk perhitungan durasi
        self.start_time_obj = datetime.now()
        # Simpan teks untuk ditampilkan di log
        self.start_time_str = self.start_time_obj.strftime("%H:%M:%S")
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        
        self.log_gui("\n=== MEMULAI PROSES PARAREL ===")
        
        # Start log listener
        self.thread_listener = threading.Thread(target=self.log_listener, daemon=True)
        self.thread_listener.start()

        # Run orchestrator in separate thread to keep UI alive
        # daemon=True agar thread mati bersama main process saat window ditutup
        t = threading.Thread(target=self.run_orchestrator, args=(file, link), daemon=True)
        t.start()

    def _kill_process_tree(self, pid):
        """Kill sebuah proses beserta seluruh child process-nya secara rekursif.
        
        Menggunakan psutil untuk menjamin browser Chrome dan chromedriver
        ikut ter-kill meskipun tidak terdaftar secara eksplisit.
        """
        if not _PSUTIL_AVAILABLE:
            # Fallback tanpa psutil: kirim SIGKILL/terminate langsung ke PID
            try:
                os.kill(pid, signal.SIGTERM)
            except (ProcessLookupError, OSError):
                pass
            return

        try:
            parent = psutil.Process(pid)
            # Kumpulkan seluruh child tree terlebih dahulu sebelum kill
            children = parent.children(recursive=True)
            # Kill dari child terdalam dulu (bottom-up)
            for child in reversed(children):
                try:
                    child.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            # Lalu kill parent process-nya
            try:
                parent.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
            # Tunggu sebentar agar OS membebaskan resource
            psutil.wait_procs([parent] + children, timeout=2)
        except (psutil.NoSuchProcess, psutil.AccessDenied, ProcessLookupError):
            pass

    def stop_automation(self):
        if not self.is_running:
            return

        self.stop_event.set()
        self.log_gui("\n[!!!] PERMINTAAN STOP DITERIMA. MENGHENTIKAN SEMUA WORKER...")

        killed_count = 0

        # Strategi 1: Kill via PID registry (Chrome + chromedriver)
        if self.worker_pids:
            for worker_id, pids in self.worker_pids.items():
                for pid in pids:
                    self.log_gui(f"   [Kill] Worker-{worker_id} | Browser PID: {pid}")
                    self._kill_process_tree(pid)
                    killed_count += 1

        # Strategi 2: Kill Python worker process (beserta child tree-nya)
        for p in self.processes:
            if p.is_alive():
                try:
                    self.log_gui(f"   [Kill] Python Worker PID: {p.pid}")
                    self._kill_process_tree(p.pid)
                    p.join(timeout=2)
                    if p.is_alive():
                        p.kill()  # Last resort
                except Exception:
                    pass
                killed_count += 1

        self.log_gui(f"[v] {killed_count} instance dihentikan paksa.")
        self.btn_stop.configure(state="disabled")
        self.is_running = False
        self.after(0, lambda: self.btn_start.configure(state="normal"))

    def log_listener(self):
        """Poll the multiprocessing queue and update the UI/Data."""
        while True:
            try:
                msg = self.queue_log.get(timeout=0.1)

                # PRIORITAS 1: Tangkap sinyal registrasi PID browser dari worker
                # Format: "[Worker-N] [PID_REGISTER]|worker_id|pid"
                # atau langsung: "[PID_REGISTER]|worker_id|pid"
                pid_match = re.search(r'\[PID_REGISTER\]\|(\d+)\|(\d+)', msg)
                if pid_match:
                    w_id = int(pid_match.group(1))
                    b_pid = int(pid_match.group(2))
                    if w_id not in self.worker_pids:
                        self.worker_pids[w_id] = []
                    self.worker_pids[w_id].append(b_pid)
                    # Tidak diteruskan ke terminal — ini sinyal internal
                    continue

                # PRIORITAS 2: Proses [ERROR_DATA] langsung di thread ini
                # agar error_data terisi sebelum orchestrator selesai.
                if "[ERROR_DATA]|" in msg:
                    # Ambil bagian setelah worker prefix jika ada
                    content = msg
                    if " [ERROR_DATA]|" in msg:
                        content = msg.split(" [ERROR_DATA]|")[1]
                        content = "[ERROR_DATA]|" + content

                    parts = content.split("|")
                    if len(parts) >= 4:
                        row_idx, nama, alasan = parts[1], parts[2], parts[3]
                        try:
                            idx = int(row_idx) - 1
                            if self.current_df is not None and 0 <= idx < len(self.current_df):
                                row_dict = self.current_df.iloc[idx].to_dict()
                                # Pastikan kunci standar ada untuk UI (biar tidak KeyError)
                                row_dict["Baris"] = row_idx
                                row_dict["Nama"] = nama
                                row_dict["Alasan"] = alasan
                                self.error_data.append(row_dict)
                        except Exception:
                            self.error_data.append({"Baris": row_idx, "Nama": nama, "Alasan": alasan})

                        # FIX: Lambda closure bug — ikat nilai ui_msg saat ini via default argument
                        # tanpa 'm=ui_msg', semua lambda akan menggunakan nilai ui_msg terakhir
                        ui_msg = f"[ERROR_UI]|{row_idx}|{nama}|{alasan}"
                        self._safe_after(0, lambda m=ui_msg: self.log_gui(m))
                else:
                    # Log biasa
                    self._safe_after(0, lambda m=msg: self.log_gui(m))

            except queue.Empty:
                # FIX: Tangkap queue.Empty secara eksplisit, bukan bare except
                # yang bisa menyembunyikan error serius seperti KeyboardInterrupt
                if not self.is_running and self.queue_log.empty():
                    break
            except Exception:
                # Guard umum agar listener thread tidak mati diam-diam
                pass

    def run_orchestrator(self, file, link):
        try:
            # 1. Read Excel
            self.current_df = pd.read_excel(file)
            df = self.current_df
            n_total = len(df)
            
            # 2. Get User Config
            headless = self.headless_var.get()
            try:
                user_workers = int(self.worker_var.get())
                n_workers = max(1, min(user_workers, self.total_cores)) # Clamp between 1 and system max
            except:
                n_workers = self.default_workers
            
            self._safe_after(0, lambda msg=f"Config: Headless={headless} | Workers={n_workers} (dari max {self.total_cores} core)": self.log_gui(msg))
            
            # 3. Chunking Data
            chunk_size = math.ceil(len(df) / n_workers)
            chunks = [df.iloc[i:i + chunk_size] for i in range(0, len(df), chunk_size)]
            
            # --- DEFINISIKAN worker_ranges DI SINI ---
            worker_ranges = []
            for chunk in chunks:
                if not chunk.empty:
                    start_idx = chunk.index[0] + 1
                    end_idx = chunk.index[-1] + 1
                    worker_ranges.append((start_idx, end_idx))
            
            self.stop_event.clear()
            self.processes = []
            
            # Merombak UI dengan mengirimkan daftar rentang baris (bukan sekadar angka jumlah worker)
            self._safe_after(0, self.setup_worker_terminals, worker_ranges)
            
            # driver_path diset ke None agar Selenium Manager yang mengelola secara native
            driver_path = None
            self._safe_after(0, lambda: self.log_gui("[System] Browser siap dijalankan via Selenium Manager."))
            
            self._safe_after(0, lambda: self.log_gui(f"Waktu Mulai: {self.start_time_str}\n----------------------------------"))
            
            # 4. Spawning Processes
            use_ts   = self.use_ts_sync_var.get()
            ss_id    = self.spreadsheet_id_var.get().strip() if use_ts else None
            creds_p  = self.creds_path_var.get().strip()    if use_ts else None

            for i, chunk in enumerate(chunks):
                if chunk.empty: continue
                p = multiprocessing.Process(
                    target=worker_launcher,
                    args=(chunk, link, self.queue_log, self.stop_event, i+1, headless, driver_path),
                    kwargs=dict(spreadsheet_id=ss_id, creds_path=creds_p)
                )
                p.start()
                self._safe_after(0, lambda msg=f"Worker-{i+1} started (PID: {p.pid})": self.log_gui(msg))
                self.processes.append(p)
                time.sleep(0.5) # Jeda 0.5 detik sesuai request
            
            # 5. Wait for all to finish & diagnostik exit code
            for p in self.processes:
                p.join()
                if p.exitcode != 0:
                    self._safe_after(0, lambda msg=f"[!] Worker PID {p.pid} exited with code {p.exitcode} (crash/error)": self.log_gui(msg))
        
        except Exception as e:
            self._safe_after(0, lambda msg=f"Orchestrator Error: {e}": self.log_gui(msg))
        
        finally:
            # Ambil waktu selesai
            end_time_obj = datetime.now()
            end_time_str = end_time_obj.strftime("%H:%M:%S")
            
            # Hitung selisih waktu (durasi)
            durasi_obj = end_time_obj - self.start_time_obj
            
            # Konversi durasi ke format 00:00:00 (Jam:Menit:Detik)
            total_detik = int(durasi_obj.total_seconds())
            jam, sisa = divmod(total_detik, 3600)
            menit, detik = divmod(sisa, 60)
            durasi_str = f"{jam:02d}:{menit:02d}:{detik:02d}"
            
            self.is_running = False
            
            # Tunggu sebentar agar listener log menghabiskan sisa queue
            if hasattr(self, 'thread_listener'):
                self.thread_listener.join(timeout=2.0)
            
            # Tambahkan jeda kecil lagi untuk memastikan semua self.after() di log_gui tuntas diproses oleh mainloop
            time.sleep(0.5)

            # Gunakan _safe_after untuk semua call UI dari thread ini
            # Mencegah RuntimeError jika window sudah ditutup saat proses masih berjalan
            self._safe_after(0, lambda: self.log_gui("\n=== SEMUA PROSES SELESAI ==="))
            
            # Menampilkan waktu
            self._safe_after(0, lambda: self.log_gui(f"Mulai   : {self.start_time_str}"))
            self._safe_after(0, lambda: self.log_gui(f"Selesai : {end_time_str}"))
            self._safe_after(0, lambda: self.log_gui(f"Durasi  : {durasi_str}"))
            
            # --- Tampilkan Panel Audit Jika Ada Error ---
            if self.error_data:
                self._safe_after(0, self.show_audit_panel)
            
            # 6. Ekspor Data Gagal ke Excel (Jika ada)
            if self.error_data:
                try:
                    folder_input = os.path.dirname(file)
                    path_export = os.path.join(folder_input, "Responden Gagal.xlsx")
                    
                    # Sort data by Baris before export
                    sorted_export = sorted(self.error_data, key=lambda x: int(x.get('Baris', 0)) if str(x.get('Baris', '')).isdigit() else 999999)
                    
                    df_err = pd.DataFrame(sorted_export)
                    df_err.to_excel(path_export, index=False)
                    self._safe_after(0, lambda: self.log_gui(f"\n[v] Berhasil mengekspor {len(self.error_data)} data gagal ke:"))
                    self._safe_after(0, lambda: self.log_gui(f"    {path_export}"))
                except Exception as ex_err:
                    self._safe_after(0, lambda e=ex_err: self.log_gui(f"\n[!] Gagal mengekspor file audit: {e}"))
            else:
                self._safe_after(0, lambda: self.log_gui("\n[i] Tidak ada responden gagal. Skip ekspor file audit."))
            
            self._safe_after(0, lambda: self.btn_start.configure(state="normal"))
            self._safe_after(0, lambda: self.btn_stop.configure(state="disabled"))

    def show_audit_panel(self):
        """Membangun panel audit di bagian bawah secara dinamis."""
        # 1. Label
        lbl_err = ctk.CTkLabel(self.frame_right, text="Ringkasan Responden Gagal (Tabel Audit)", font=ctk.CTkFont(size=12, weight="bold"), text_color="#ff5555")
        lbl_err.grid(row=4, column=0, columnspan=10, sticky="w", padx=10, pady=(20, 0))
        
        # 2. Textbox
        self.error_terminal = ctk.CTkTextbox(self.frame_right, height=120, font=("Consolas", 10), border_width=1, border_color="#ff5555")
        self.error_terminal.grid(row=5, column=0, columnspan=10, sticky="ew", padx=10, pady=(0, 10))
        
        # 3. Isi Data
        content = "Baris | Nama | Alasan\n" + "-"*50 + "\n"
        
        # Sort data by Baris (numeric) before display
        sorted_errors = sorted(self.error_data, key=lambda x: int(x.get('Baris', 0)) if str(x.get('Baris', '')).isdigit() else 999999)
        
        for err in sorted_errors:
            # Gunakan .get() sebagai pengaman extra
            b = str(err.get('Baris', '-'))
            n = str(err.get('Nama', 'Unknown'))
            a = str(err.get('Alasan', 'Error'))
            content += f"{b.ljust(6)} | {n.ljust(20)} | {a}\n"
        
        self.error_terminal.insert("0.0", content)
        self.error_terminal.configure(state="disabled")
        self.error_terminal.see("end")
