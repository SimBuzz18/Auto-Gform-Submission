import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import os
import threading
from AutoForm import AutoFormLogic

# --- GUI APP ---
class AutoFormApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Window Config
        self.title("AutoForm Automation V1.0")
        self.geometry("900x500")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Layout Config
        self.grid_columnconfigure(1, weight=1) # Kanan expand
        self.grid_rowconfigure(0, weight=1)    # Full height

        # State Variables
        self.is_running = False
        self.stop_requested = False
        self.link_var = tk.StringVar()
        self.file_path_var = tk.StringVar()
        
        # --- LEFT PANEL (INPUTS) ---
        self.frame_left = ctk.CTkFrame(self, width=300, corner_radius=0)
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
        lbl_title = ctk.CTkLabel(self.frame_left, text="Aplikasi Otomatisasi v2.0", font=ctk.CTkFont(size=20, weight="bold"))
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

        # Control Buttons
        self.btn_start = ctk.CTkButton(self.frame_left, text="START", command=self.start_automation, height=40, font=ctk.CTkFont(weight="bold"))
        self.btn_start.grid(row=6, column=0, padx=20, pady=(30, 10), sticky="ew")
        
        self.btn_stop = ctk.CTkButton(self.frame_left, text="STOP", command=self.stop_automation, height=40, fg_color="darkred", hover_color="#aa0000", state="disabled")
        self.btn_stop.grid(row=7, column=0, padx=20, pady=10, sticky="ew")

    def _init_right_panel(self):
        lbl_term = ctk.CTkLabel(self.frame_right, text="Terminal Output & Log", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_term.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        # Textbox (Terminal)
        self.terminal = ctk.CTkTextbox(self.frame_right, width=400, font=("Consolas", 12))
        self.terminal.grid(row=1, column=0, padx=10, pady=10, sticky="nswe")
        self.terminal.insert("0.0", "Siap menjalankan otomatisasi...\n")
        self.terminal.configure(state="disabled") # Readonly

    def browse_file(self):
        self.attributes('-topmost', False)
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        self.attributes('-topmost', True)
        if filename:
            self.file_path_var.set(filename)
            self.log_gui(f"File dipilih: {os.path.basename(filename)}")

    def log_gui(self, message):
        """Append text to the GUI terminal in a thread-safe way."""
        self.terminal.configure(state="normal")
        self.terminal.insert("end", f"{message}\n")
        self.terminal.see("end")
        self.terminal.configure(state="disabled")

    def start_automation(self):
        if self.is_running: return
        
        link = self.link_var.get().strip()
        file = self.file_path_var.get().strip()
        
        if not link or not file:
            tk.messagebox.showwarning("Warning", "Link dan File Excel harus diisi!")
            return

        self.is_running = True
        self.stop_requested = False
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        
        self.log_gui("\n=== MEMULAI PROSES ===")
        
        # Run in separate thread
        t = threading.Thread(target=self.run_logic_thread, args=(file, link))
        t.start()

    def stop_automation(self):
        if self.is_running:
            self.stop_requested = True
            self.log_gui("\n[!!!] PERMINTAAN STOP DITERIMA. SEDANG MENYELESAIKAN PROSES SAAT INI...")
            self.btn_stop.configure(state="disabled") # Prevent double click

    def check_stop(self):
        return self.stop_requested

    def run_logic_thread(self, file, link):
        # Instantiate Logic with callbacks
        logic = AutoFormLogic(self.log_gui, self.check_stop)
        logic.run_process(file, link)
        
        self.is_running = False
        self.stop_requested = False
        
        # Reset UI
        self.after(0, lambda: self.btn_start.configure(state="normal"))
        self.after(0, lambda: self.btn_stop.configure(state="disabled"))


if __name__ == "__main__":
    app = AutoFormApp()
    # Force top level
    app.attributes('-topmost', True)
    app.mainloop()
