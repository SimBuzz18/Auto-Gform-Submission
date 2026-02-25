import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import os
import time
import threading
import multiprocessing
import pandas as pd
import math
import re
from datetime import datetime
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
        self.queue_log = multiprocessing.Queue()
        self.stop_event = multiprocessing.Event()
        self.link_var = tk.StringVar()
        self.file_path_var = tk.StringVar()
        self.error_data = [] # Koleksi data responden gagal: {'Baris': x, 'Nama': y, 'Alasan': z}
        
        # New Settings Variables
        self.headless_var = tk.BooleanVar(value=True)
        self.total_cores = multiprocessing.cpu_count()
        self.default_workers = max(1, int(self.total_cores * 0.75))
        self.worker_var = tk.StringVar(value=str(self.default_workers))
        self.worker_var.trace_add("write", self._validate_worker_range)
        
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
        self.btn_start = ctk.CTkButton(self.frame_left, text="START", command=self.start_automation, height=40, font=ctk.CTkFont(weight="bold"))
        self.btn_start.grid(row=7, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        self.btn_stop = ctk.CTkButton(self.frame_left, text="STOP", command=self.stop_automation, height=40, fg_color="darkred", hover_color="#aa0000", state="disabled")
        self.btn_stop.grid(row=8, column=0, padx=20, pady=10, sticky="ew")

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
        self.attributes('-topmost', False)
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        self.attributes('-topmost', True)
        if filename:
            self.file_path_var.set(filename)
            self.log_gui(f"File dipilih: {os.path.basename(filename)}")

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

        # B. Cek apakah teks mengandung identitas worker

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
        self.stop_requested = False
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
        t = threading.Thread(target=self.run_orchestrator, args=(file, link))
        t.start()

    def stop_automation(self):
        if self.is_running:
            self.stop_event.set()
            self.log_gui("\n[!!!] PERMINTAAN STOP DITERIMA. MENGHENTIKAN SEMUA WORKER...")
            self.btn_stop.configure(state="disabled")

    def log_listener(self):
        """Poll the multiprocessing queue and update the UI/Data."""
        while True:
            try:
                msg = self.queue_log.get(timeout=0.1)
                
                # SINKRONISASI DATA: Proses [ERROR_DATA] langsung di thread ini
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
                        except:
                            self.error_data.append({"Baris": row_idx, "Nama": nama, "Alasan": alasan})
                        
                        # Teruskan ke UI untuk update terminal error (jika panel sudah dibuka)
                        ui_msg = f"[ERROR_UI]|{row_idx}|{nama}|{alasan}"
                        self.after(0, lambda: self.log_gui(ui_msg))
                else:
                    # Log biasa
                    self.after(0, lambda m=msg: self.log_gui(m))
                    
            except:
                if not self.is_running and self.queue_log.empty(): 
                    break

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
            
            self.log_gui(f"Config: Headless={headless} | Workers={n_workers} (dari max {self.total_cores} core)")
            
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
            self.after(0, self.setup_worker_terminals, worker_ranges)
            
            # Download & Cache driver 1 kali sebelum worker berjalan agar tidak tabrakan
            ChromeDriverManager().install()
            
            # 4. Spawning Processes
            for i, chunk in enumerate(chunks):
                if chunk.empty: continue
                p = multiprocessing.Process(
                    target=worker_launcher, 
                    args=(chunk, link, self.queue_log, self.stop_event, i+1, headless)
                )
                p.start()
                self.processes.append(p)
            
            # 5. Wait for all to finish
            for p in self.processes:
                p.join()
        
        except Exception as e:
            self.log_gui(f"Orchestrator Error: {e}")
        
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

            self.log_gui("\n=== SEMUA PROSES SELESAI ===")
            
            # Menampilkan waktu dengan spasi agar tanda ':' sejajar vertikal
            self.log_gui(f"Mulai   : {self.start_time_str}")
            self.log_gui(f"Selesai : {end_time_str}")
            self.log_gui(f"Durasi  : {durasi_str}")
            
            # --- Tampilkan Panel Audit Jika Ada Error ---
            if self.error_data:
                self.after(0, self.show_audit_panel)
            
            # 6. Ekspor Data Gagal ke Excel (Jika ada)
            if self.error_data:
                try:
                    folder_input = os.path.dirname(file)
                    path_export = os.path.join(folder_input, "Responden Gagal.xlsx")
                    
                    # Sort data by Baris before export
                    sorted_export = sorted(self.error_data, key=lambda x: int(x.get('Baris', 0)) if str(x.get('Baris', '')).isdigit() else 999999)
                    
                    df_err = pd.DataFrame(sorted_export)
                    df_err.to_excel(path_export, index=False)
                    self.log_gui(f"\n[v] Berhasil mengekspor {len(self.error_data)} data gagal ke:")
                    self.log_gui(f"    {path_export}")
                except Exception as ex_err:
                    self.log_gui(f"\n[!] Gagal mengekspor file audit: {ex_err}")
            else:
                self.log_gui("\n[i] Tidak ada responden gagal. Skip ekspor file audit.")
            
            self.after(0, lambda: self.btn_start.configure(state="normal"))
            self.after(0, lambda: self.btn_stop.configure(state="disabled"))

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


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = AutoFormApp()
    # Force top level
    app.attributes('-topmost', True)
    app.mainloop()
