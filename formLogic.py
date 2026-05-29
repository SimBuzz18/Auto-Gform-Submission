import pandas as pd
import time
import re
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
)
from webdriver_manager.chrome import ChromeDriverManager

# Optional: Google Sheets API untuk patch timestamp setelah submit
try:
    import gspread
    from google.oauth2.service_account import Credentials as GSCredentials
    _GSPREAD_AVAILABLE = True
except ImportError:
    _GSPREAD_AVAILABLE = False

# Nama kolom timestamp yang dikenal (case-insensitive via normalize_text)
_TIMESTAMP_COL_ALIASES = {
    'timestamp', 'stempel waktu', 'waktu', 'waktu submit',
    'tanggal submit', 'tanggal', 'time', 'submit time', 'submit at',
}


class SkipRespondentException(Exception):
    """Exception raised to skip current respondent and move to next."""
    pass


def worker_launcher(
    df_chunk, link_gform, queue_log, stop_event, worker_id,
    headless=True, driver_path=None,
    spreadsheet_id=None, creds_path=None
):
    """Entry point for multiprocessing worker.

    PENTING: Fungsi ini adalah titik masuk pertama yang dieksekusi oleh child process
    setelah spawn. Jika terjadi crash di sini (import error, pickle error, dsb),
    error HARUS dikirim ke queue_log agar muncul di terminal GUI.
    Tanpa try/except ini, error hanya masuk ke stderr yang bisa jadi /dev/null
    di PyInstaller --noconsole mode → crash diam-diam, worker tampak tidak jalan.
    """
    try:
        form_logic = logic(
            callback_log=queue_log,
            stop_event=stop_event,
            worker_id=worker_id,
            spreadsheet_id=spreadsheet_id,
            creds_path=creds_path,
        )
        form_logic.run_process_chunk(df_chunk, link_gform, headless=headless, driver_path=driver_path)
    except Exception as e:
        import traceback
        error_msg = f"[Worker-{worker_id}] [FATAL] Worker crash: {e}\n{traceback.format_exc()}"
        if queue_log and hasattr(queue_log, "put"):
            queue_log.put(error_msg)
        else:
            print(error_msg)


class logic:
    def __init__(
        self,
        callback_log=None,
        stop_event=None,
        worker_id=None,
        spreadsheet_id=None,
        creds_path=None,
    ):
        """
        callback_log    : fungsi atau multiprocessing.Queue untuk logging
        stop_event      : multiprocessing.Event | None
        worker_id       : identifier logging
        spreadsheet_id  : ID Google Spreadsheet respons form (opsional)
        creds_path      : path ke service account JSON Google (opsional)
        """
        self.callback_log    = callback_log
        self.stop_event      = stop_event
        self.worker_id       = worker_id
        self.separator_checkbox = ','
        self.spreadsheet_id  = spreadsheet_id
        self.creds_path      = creds_path
        self._gsheet         = None  # Lazy-init, reused sepanjang satu chunk

    def _log(self, message):
        prefix = f"[Worker-{self.worker_id}] " if self.worker_id is not None else ""
        msg = f"{prefix}{message}"
        if self.callback_log:
            if hasattr(self.callback_log, "put"): # It's a Queue
                self.callback_log.put(msg)
            else: # It's a function
                self.callback_log(msg)
        else:
            print(msg)

    def _is_stopped(self):
        if self.stop_event:
            return self.stop_event.is_set()
        return False

    def clean_data(self, val):
        val = str(val).strip()
        if val.endswith('.0'):
            return val[:-2]
        return val

    def normalize_text(self, text):
        if not isinstance(text, str): return ""
        return re.sub(r'\s+', ' ', text.lower().strip())

    def _click_element(self, driver, element):
        """Klik elemen menggunakan metode hibrida: native click dengan fallback ke JS click."""
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            time.sleep(0.15)
            element.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", element)
            except Exception as e:
                self._log(f"   [!] Gagal mengklik elemen: {e}")
                raise

    # =========================================================
    # TIMESTAMP SYNC — Google Sheets API
    # =========================================================

    def _format_timestamp(self, raw) -> str:
        """Normalisasi berbagai format timestamp ke format Google Sheets: DD/MM/YYYY HH:MM:SS.

        Mendukung: pandas Timestamp, datetime, string ISO, string Indonesian.
        """
        try:
            if hasattr(raw, 'to_pydatetime'):  # pandas Timestamp
                dt = raw.to_pydatetime()
            elif isinstance(raw, datetime):
                dt = raw
            else:
                raw_str = str(raw).strip()
                dt = None
                for fmt in (
                    '%d/%m/%Y %H:%M:%S',
                    '%Y-%m-%d %H:%M:%S',
                    '%d/%m/%Y %H:%M',
                    '%Y-%m-%d %H:%M',
                    '%d-%m-%Y %H:%M:%S',
                    '%d/%m/%Y',
                    '%Y-%m-%d',
                ):
                    try:
                        dt = datetime.strptime(raw_str, fmt)
                        break
                    except ValueError:
                        continue
                if dt is None:
                    # Fallback: serahkan ke pandas parser
                    dt = pd.to_datetime(raw_str).to_pydatetime()
            return dt.strftime('%d/%m/%Y %H:%M:%S')
        except Exception:
            return str(raw)  # Kembalikan as-is jika semua parser gagal

    def _get_sheet(self):
        """Lazy-init koneksi ke sheet pertama Google Spreadsheet.

        Raises RuntimeError jika dependency atau konfigurasi tidak tersedia.
        """
        if self._gsheet is not None:
            return self._gsheet
        if not _GSPREAD_AVAILABLE:
            raise RuntimeError(
                "Library gspread belum terinstall. Jalankan: pip install gspread google-auth"
            )
        if not self.creds_path or not self.spreadsheet_id:
            raise RuntimeError(
                "spreadsheet_id dan creds_path wajib diisi untuk fitur sinkronisasi timestamp."
            )
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive',
        ]
        creds = GSCredentials.from_service_account_file(self.creds_path, scopes=scopes)
        client = gspread.authorize(creds)
        self._gsheet = client.open_by_key(self.spreadsheet_id).sheet1
        self._log("[TS] Koneksi Google Sheets berhasil.")
        return self._gsheet

    def _patch_timestamp(self, sheet, target_ts_str: str, row_before_submit: int):
        """Update cell Timestamp (kolom A) di baris respons yang baru saja masuk.

        Strategi: Bandingkan jumlah baris sebelum dan sesudah submit.
        Jika ada baris baru → update baris terakhir.
        Timeout 10 detik menunggu Google memproses respons.
        """
        deadline = time.time() + 10
        while time.time() < deadline:
            try:
                all_vals = sheet.get_all_values()
                current_rows = len(all_vals)
                if current_rows > row_before_submit:
                    target_row = current_rows  # 1-indexed, header di baris 1
                    sheet.update_cell(target_row, 1, target_ts_str)
                    self._log(f"   [v] Timestamp di-patch → {target_ts_str} (baris {target_row})")
                    return
                time.sleep(1.5)  # Tunggu Google Sheets menerima respons
            except Exception as e:
                self._log(f"   [!] Gagal patch timestamp (retry): {e}")
                time.sleep(2)
        self._log("   [!] Timeout menunggu baris baru di Spreadsheet — timestamp tidak di-patch.")

    def run_process_chunk(self, df_chunk, link_gform, headless=False, driver_path=None):
        """Runs the automation for a specific chunk of data."""
        driver = None
        try:
            # Normalize column names
            col_map = {self.normalize_text(col): col for col in df_chunk.columns}
            
            # Setup Browser
            self._log("Membuka Browser Chrome...")
            options = webdriver.ChromeOptions()
            if headless:
                options.add_argument("--headless=new")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--disable-gpu")
                options.add_argument("--window-size=1920,1080")
            
            if driver_path:
                driver = webdriver.Chrome(service=Service(driver_path), options=options)
            else:
                driver = webdriver.Chrome(options=options)

            # Daftarkan PID browser ke GUI agar bisa di-kill secara eksplisit
            # saat tombol STOP ditekan (tanpa bergantung pada proses Python wrapper)
            try:
                browser_pid = driver.service.process.pid
                self._log(f"[PID_REGISTER]|{self.worker_id}|{browser_pid}")
            except Exception:
                pass  # Jika gagal ambil PID, abaikan — fallback ke kill via Python process tree

            MAX_RETRIES = 3  # Maksimal percobaan per responden

            # Deteksi kolom timestamp dari header Excel (sekali sebelum loop)
            ts_col_key = None
            for col_key in col_map:
                if col_key in _TIMESTAMP_COL_ALIASES:
                    ts_col_key = col_key
                    self._log(f"[TS] Kolom timestamp terdeteksi: '{col_map[ts_col_key]}'")
                    break

            # Inisialisasi koneksi Sheets sekali sebelum loop (reused per worker)
            sheet = None
            _use_ts_sync = bool(self.spreadsheet_id and self.creds_path)
            if _use_ts_sync:
                try:
                    sheet = self._get_sheet()
                except Exception as e:
                    self._log(f"[TS] [!] Gagal init koneksi Sheets: {e} — fitur timestamp dinonaktifkan.")
                    _use_ts_sync = False

            for index, row in df_chunk.iterrows():
                if self._is_stopped():
                    self._log("[STOP] Proses dihentikan.")
                    break

                # FIX #1: Ambil nama_resp sekali sebelum retry loop
                # agar tidak UnboundLocalError jika stop_event langsung aktif
                # FIX #2: pandas Series.get() bersarang tidak menghasilkan fallback yang benar;
                # gunakan 'or' chain agar fallback berjalan secara lazy
                nama_resp = row.get('Nama Lengkap') or row.get('Nama') or 'Unknown'

                # WebDriverWait dibuat di luar retry loop (tidak perlu re-instantiate tiap percobaan)
                wait = WebDriverWait(driver, 10)

                for attempt in range(1, MAX_RETRIES + 1):
                    if self._is_stopped(): break
                    
                    if attempt == 1:
                        self._log(f"Resp. ke-{index + 1}...")
                    else:
                        self._log(f"   [!] Retrying Resp. ke-{index + 1} (Percobaan {attempt}/{MAX_RETRIES})...")

                    try:
                        driver.get(link_gform)
                        
                        # Menunggu hingga dokumen browser selesai terload penuh (100%)
                        try:
                            WebDriverWait(driver, 30).until(
                                lambda d: d.execute_script("return document.readyState") == "complete"
                            )
                            self._log("   [v] Browser selesai dimuat (100% terload).")
                        except TimeoutException:
                            self._log("   [!] Timeout menunggu document.readyState == 'complete'")
                        
                        halaman_ke = 1
                        sukses_kirim = False # Penanda apakah berhasil sampai tombol Kirim
                        
                        while True:
                            if self._is_stopped(): break
                            
                            # Update: Use dynamic wait and stabilization loop for questions
                            try:
                                wait.until(EC.presence_of_element_located((By.XPATH, "//div[@role='listitem']")))
                            except TimeoutException:
                                self._log("   [!] Timeout menunggu pertanyaan muncul.")
                                break

                            # Loop Stabilisasi: Tunggu hingga jumlah elemen pertanyaan stabil (tidak berubah selama 1 detik)
                            last_count = -1
                            stable_time = time.time()
                            timeout_stable = time.time() + 10
                            while time.time() < timeout_stable:
                                if self._is_stopped(): break
                                current_count = len(driver.find_elements(By.XPATH, "//div[@role='listitem']"))
                                if current_count != last_count:
                                    last_count = current_count
                                    stable_time = time.time()
                                elif time.time() - stable_time >= 1.0:
                                    break
                                time.sleep(0.2)

                            questions = driver.find_elements(By.XPATH, "//div[@role='listitem']")
                            
                            # (Loop pengisian pertanyaan tetap sama seperti sebelumnya)
                            for q_elem in questions:
                                try:
                                    try:
                                        heading_elem = q_elem.find_element(By.XPATH, ".//div[@role='heading']/span")
                                        q_text = heading_elem.text
                                    except NoSuchElementException:
                                        # Elemen bukan pertanyaan (misalnya section header), lewati
                                        continue

                                    normalized_q_text = self.normalize_text(q_text)
                                    target_col = None
                                    
                                    if normalized_q_text in col_map:
                                        target_col = col_map[normalized_q_text]
                                    
                                    # FIX #3: Gunakan 'is None' bukan 'not' agar kolom bernama
                                    # '0', ' ', atau falsy string lainnya tidak dianggap tidak ditemukan
                                    if target_col is None:
                                        raise SkipRespondentException(f"Pertanyaan '{q_text}' tidak ditemukan di Excel")
                                        
                                    answer = row[target_col]
                                    if pd.isna(answer) or str(answer).lower() == 'nan' or str(answer).strip() == '':
                                        continue 

                                    # =========================================================
                                    # HANDLER URUTAN DETEKSI TIPE INPUT GOOGLE FORM
                                    # Urutan penting: Grid harus dicek SEBELUM Radio/Checkbox
                                    # biasa karena keduanya mengandung elemen yang sama.
                                    # =========================================================

                                    # A. Kisi-Kisi Pilihan Ganda (Multiple Choice Grid)
                                    # Deteksi: > 1 div[role='radiogroup'] di dalam q_elem
                                    # (tiap radiogroup = 1 baris di grid)
                                    # Format Excel: "NamaBaris1:JawabanKolom;NamaBaris2:JawabanKolom"
                                    radiogroups = q_elem.find_elements(By.XPATH, ".//div[@role='radiogroup']")
                                    if len(radiogroups) > 1:
                                        # Parse format "Baris1:Kolom1;Baris2:Kolom2"
                                        row_answers = {}
                                        for pair in str(answer).split(';'):
                                            pair = pair.strip()
                                            if ':' in pair:
                                                r_name, r_val = pair.split(':', 1)
                                                row_answers[self.normalize_text(r_name.strip())] = r_val.strip()

                                        if not row_answers:
                                            raise SkipRespondentException(
                                                f"Format grid '{q_text}' salah. Gunakan: 'Baris1:Jawaban;Baris2:Jawaban'"
                                            )

                                        for rg in radiogroups:
                                            rg_label = self.normalize_text(rg.get_attribute('aria-label') or '')
                                            if rg_label not in row_answers:
                                                continue  # Baris tidak ada di Excel, lewati

                                            target_val = row_answers[rg_label]
                                            radios_in_row = rg.find_elements(By.XPATH, ".//div[@role='radio']")
                                            found = False
                                            for radio in radios_in_row:
                                                dv = radio.get_attribute('data-value') or ''
                                                al = radio.get_attribute('aria-label') or ''
                                                if (self.normalize_text(dv) == self.normalize_text(target_val) or
                                                        self.normalize_text(al) == self.normalize_text(target_val)):
                                                    if radio.get_attribute('aria-checked') == 'false':
                                                        self._click_element(driver, radio)
                                                    found = True
                                                    break
                                            if not found:
                                                raise SkipRespondentException(
                                                    f"Pilihan '{target_val}' tidak ada di baris '{rg_label}' — grid '{q_text}'"
                                                )
                                        continue

                                    # B. Petak Kotak Centang (Checkbox Grid)
                                    # Deteksi: checkbox di mana aria-label mengandung pola "NamaBaris, NamaKolom"
                                    # Format Excel: "Baris1:Kolom1|Kolom2;Baris2:Kolom3"
                                    all_checkboxes = q_elem.find_elements(By.XPATH, ".//div[@role='checkbox']")
                                    if all_checkboxes:
                                        # Sampling aria-label untuk mendeteksi apakah ini checkbox grid
                                        sample_labels = [
                                            (cb.get_attribute('aria-label') or '') for cb in all_checkboxes[:6]
                                        ]
                                        # Grid: aria-label format = "NamaBaris, NamaKolom"
                                        row_prefixes = set()
                                        for lbl in sample_labels:
                                            parts = lbl.rsplit(',', 1)
                                            if len(parts) == 2:
                                                row_prefixes.add(parts[0].strip().lower())
                                        is_checkbox_grid = len(row_prefixes) > 1

                                        if is_checkbox_grid:
                                            # Parse format "Baris1:Kolom1|Kolom2;Baris2:Kolom3"
                                            row_answers = {}
                                            for pair in str(answer).split(';'):
                                                pair = pair.strip()
                                                if ':' in pair:
                                                    r_name, r_vals = pair.split(':', 1)
                                                    col_list = [self.normalize_text(v) for v in r_vals.split('|')]
                                                    row_answers[self.normalize_text(r_name.strip())] = col_list

                                            if not row_answers:
                                                raise SkipRespondentException(
                                                    f"Format checkbox grid '{q_text}' salah. Gunakan: 'Baris1:Kolom1|Kolom2;Baris2:Kolom3'"
                                                )

                                            for cb in all_checkboxes:
                                                try:
                                                    aria_lbl = cb.get_attribute('aria-label') or ''
                                                    parts = aria_lbl.rsplit(',', 1)
                                                    if len(parts) != 2:
                                                        continue
                                                    row_norm = self.normalize_text(parts[0].strip())
                                                    col_norm = self.normalize_text(parts[1].strip())

                                                    if row_norm in row_answers and col_norm in row_answers[row_norm]:
                                                        if cb.get_attribute('aria-checked') == 'false':
                                                            self._click_element(driver, cb)
                                                except StaleElementReferenceException:
                                                    continue
                                            continue

                                        # Bukan grid — proses sebagai Checkbox biasa (Section D)
                                        answers_list = [self.normalize_text(x) for x in str(answer).split(self.separator_checkbox)]
                                        found_count = 0
                                        for cb in all_checkboxes:
                                            data_val = cb.get_attribute("data-value")
                                            aria_label = cb.get_attribute("aria-label")
                                            val_norm = ""
                                            if data_val: val_norm = self.normalize_text(data_val)
                                            elif aria_label: val_norm = self.normalize_text(aria_label)
                                            if val_norm in answers_list:
                                                if cb.get_attribute("aria-checked") == "false":
                                                    self._click_element(driver, cb)
                                                found_count += 1
                                        if found_count < len(answers_list):
                                            raise SkipRespondentException(f"Opsi checkbox '{answer}' tidak lengkap di Form")
                                        continue

                                    # C. Radio Button (termasuk Skala Linear & Rating)
                                    # Skala Linear & Rating di GForm menggunakan div[role='radio']
                                    # dengan data-value berupa angka — handler ini sudah menanganinya.
                                    # Format Excel Skala Linear: nilai numerik (misal: "3")
                                    radios = q_elem.find_elements(By.XPATH, ".//div[@role='radio']")
                                    if radios:
                                        found_radio = False
                                        for radio in radios:
                                            data_val = radio.get_attribute("data-value")
                                            aria_label = radio.get_attribute("aria-label")
                                            is_match = False
                                            if data_val and self.normalize_text(data_val) == self.normalize_text(str(answer)): is_match = True
                                            elif aria_label and self.normalize_text(aria_label) == self.normalize_text(str(answer)): is_match = True
                                            if is_match:
                                                if radio.get_attribute("aria-checked") == "false":
                                                    self._click_element(driver, radio)
                                                found_radio = True
                                                break
                                        if not found_radio:
                                            raise SkipRespondentException(f"Pilihan '{answer}' tidak ada untuk pertanyaan '{q_text}'")
                                        continue

                                    # D. Dropdown (Google Form custom listbox — BUKAN <select> native)
                                    # Struktur: div[role='listbox'] sebagai trigger klik,
                                    # lalu opsi muncul di div[role='option'] setelah popup terbuka.
                                    dropdown_trigger = q_elem.find_elements(
                                        By.XPATH, ".//div[@role='listbox']"
                                    )
                                    if dropdown_trigger:
                                        trigger = dropdown_trigger[0]
                                        try:
                                            self._click_element(driver, trigger)
                                        except Exception:
                                            trigger.click()

                                        # Tunggu popup opsi muncul di DOM (lazy render)
                                        try:
                                            wait_dd = WebDriverWait(driver, 5)
                                            wait_dd.until(
                                                EC.presence_of_element_located(
                                                    (By.XPATH, "//div[@role='option'] | //li[@role='option']")
                                                )
                                            )
                                        except TimeoutException:
                                            raise SkipRespondentException(
                                                f"Dropdown '{q_text}' tidak membuka opsi (timeout)"
                                            )

                                        options = driver.find_elements(
                                            By.XPATH,
                                            "//div[@role='option'] | //li[@role='option']"
                                        )
                                        found_option = False
                                        for opt in options:
                                            try:
                                                opt_text = opt.get_attribute("data-value") or opt.text
                                                if self.normalize_text(opt_text) == self.normalize_text(str(answer)):
                                                    self._click_element(driver, opt)
                                                    found_option = True
                                                    break
                                            except StaleElementReferenceException:
                                                continue

                                        if not found_option:
                                            try:
                                                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                                            except Exception:
                                                pass
                                            raise SkipRespondentException(
                                                f"Pilihan dropdown '{answer}' tidak ditemukan di '{q_text}'"
                                            )
                                        continue

                                    # E. Tanggal (Date)
                                    # GForm merender 3 input terpisah: Day, Month, Year
                                    # Format Excel: "DD/MM/YYYY" atau objek datetime pandas
                                    date_day = q_elem.find_elements(By.XPATH, ".//input[@aria-label='Day' or @placeholder='DD']")
                                    if not date_day:
                                        date_day = q_elem.find_elements(By.XPATH, ".//input[@type='date']")
                                    if date_day:
                                        try:
                                            # Normalisasi: datetime pandas, string DD/MM/YYYY, atau YYYY-MM-DD
                                            raw_val = str(answer).strip()
                                            dt = None
                                            for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%m/%d/%Y'):
                                                try:
                                                    dt = datetime.strptime(raw_val, fmt)
                                                    break
                                                except ValueError:
                                                    continue
                                            if dt is None:
                                                # Coba parse dari format datetime pandas (Timestamp)
                                                import pandas as _pd
                                                dt = _pd.to_datetime(raw_val).to_pydatetime()

                                            # Isi input terpisah Day / Month / Year
                                            inp_day   = q_elem.find_elements(By.XPATH, ".//input[@aria-label='Day' or @placeholder='DD']")
                                            inp_month = q_elem.find_elements(By.XPATH, ".//input[@aria-label='Month' or @placeholder='MM']")
                                            inp_year  = q_elem.find_elements(By.XPATH, ".//input[@aria-label='Year' or @placeholder='YYYY']")

                                            def _fill_date_input(elems, val):
                                                if elems:
                                                    el = elems[0]
                                                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                                                    el.clear()
                                                    el.send_keys(str(val))

                                            if inp_day and inp_month and inp_year:
                                                _fill_date_input(inp_day,   dt.day)
                                                _fill_date_input(inp_month, dt.month)
                                                _fill_date_input(inp_year,  dt.year)
                                            else:
                                                # Fallback: single date input
                                                el = date_day[0]
                                                el.clear()
                                                el.send_keys(dt.strftime('%d/%m/%Y'))
                                            time.sleep(0.2)
                                        except Exception as de:
                                            raise SkipRespondentException(
                                                f"Gagal mengisi tanggal '{answer}' di '{q_text}': {de}"
                                            )
                                        continue

                                    # F. Waktu (Time)
                                    # GForm merender 2 input terpisah: Hour, Minute
                                    # Format Excel: "HH:MM" (24 jam)
                                    time_hour = q_elem.find_elements(By.XPATH, ".//input[@aria-label='Hour' or @placeholder='HH']")
                                    if not time_hour:
                                        time_hour = q_elem.find_elements(By.XPATH, ".//input[@type='time']")
                                    if time_hour:
                                        try:
                                            raw_val = str(answer).strip()
                                            # Normalisasi format: "HH:MM" atau "H:MM"
                                            parts = raw_val.replace('.', ':').split(':')
                                            hh = int(parts[0]) if len(parts) >= 1 else 0
                                            mm = int(parts[1]) if len(parts) >= 2 else 0

                                            inp_hour   = q_elem.find_elements(By.XPATH, ".//input[@aria-label='Hour' or @placeholder='HH']")
                                            inp_minute = q_elem.find_elements(By.XPATH, ".//input[@aria-label='Minute' or @placeholder='MM']")

                                            def _fill_time_input(elems, val):
                                                if elems:
                                                    el = elems[0]
                                                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                                                    el.clear()
                                                    el.send_keys(f"{val:02d}")

                                            if inp_hour and inp_minute:
                                                _fill_time_input(inp_hour,   hh)
                                                _fill_time_input(inp_minute, mm)
                                            else:
                                                # Fallback: single time input
                                                el = time_hour[0]
                                                el.clear()
                                                el.send_keys(f"{hh:02d}:{mm:02d}")
                                            time.sleep(0.2)
                                        except Exception as te:
                                            raise SkipRespondentException(
                                                f"Gagal mengisi waktu '{answer}' di '{q_text}': {te}"
                                            )
                                        continue

                                    # G. Text Input / Textarea (fallback terakhir)
                                    text_inputs = q_elem.find_elements(By.XPATH, ".//input[@type='text'] | .//textarea")
                                    visible_inputs = [inp for inp in text_inputs if inp.is_displayed()]

                                    if visible_inputs:
                                        inp = visible_inputs[0]
                                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", inp)
                                        time.sleep(0.3)
                                        try:
                                            inp.clear()
                                        except Exception:
                                            pass
                                        try:
                                            inp.send_keys(str(answer))
                                        except Exception:
                                            driver.execute_script("arguments[0].value = arguments[1];", inp, str(answer))
                                            driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", inp)
                                            driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", inp)
                                        continue

                                except SkipRespondentException:
                                    raise  # Re-raise to skip respondent
                                except Exception as e:
                                    raise SkipRespondentException(f"Error Pertanyaan: {str(e)}")

                            # --- NAVIGASI ---
                            if self._is_stopped(): break

                            try:
                                # Check for Submit or Next button with dynamic wait (short timeout for polling)
                                wait_nav = WebDriverWait(driver, 3)
                                nav_btn = wait_nav.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='button']//span[text()='Kirim' or text()='Submit' or text()='Berikutnya' or text()='Next']")))
                                btn_text = nav_btn.text.lower()
                                
                                if "kirim" in btn_text or "submit" in btn_text:
                                    driver.execute_script("arguments[0].scrollIntoView();", nav_btn)
                                    time.sleep(0.3)

                                    # Capture jumlah baris SEBELUM submit untuk timestamp patching
                                    row_before_submit = None
                                    if _use_ts_sync and sheet is not None:
                                        try:
                                            row_before_submit = len(sheet.get_all_values())
                                        except Exception as _e:
                                            self._log(f"   [TS] Tidak dapat baca Sheets sebelum submit: {_e}")

                                    self._click_element(driver, nav_btn)
                                    
                                    # Tunggu hingga halaman konfirmasi terload 100% dan respon tercatat
                                    self._log("   Menunggu konfirmasi pengiriman dari Google Form...")
                                    confirm_deadline = time.time() + 15
                                    while time.time() < confirm_deadline:
                                        if self._is_stopped(): break
                                        
                                        current_url = driver.current_url.lower()
                                        page_source = driver.page_source.lower()
                                        
                                        # Indikator sukses: URL mengandung 'formresponse' ATAU ada teks konfirmasi (Indonesian/English)
                                        has_confirm_text = any(txt in page_source for txt in [
                                            "tanggapan anda telah direkam",
                                            "your response has been recorded",
                                            "kirim tanggapan lain",
                                            "submit another response",
                                            "edit tanggapan",
                                            "edit your response"
                                        ])
                                        
                                        # Cek apakah elemen pertanyaan masih ada
                                        questions_left = driver.find_elements(By.XPATH, "//div[@role='listitem']")
                                        
                                        if ("formresponse" in current_url or has_confirm_text) and len(questions_left) == 0:
                                            sukses_kirim = True
                                            break
                                            
                                        time.sleep(0.5)

                                    if sukses_kirim:
                                        self._log(f"[v] Resp. {index + 1} TERKIRIM.")
                                        
                                        # Patch timestamp segera setelah submit berhasil
                                        if _use_ts_sync and sheet is not None and row_before_submit is not None:
                                            # Ambil nilai timestamp dari Excel, atau fallback ke waktu sekarang
                                            ts_raw = None
                                            if ts_col_key is not None:
                                                ts_raw = row.get(col_map[ts_col_key])

                                            if ts_raw is not None and not pd.isna(ts_raw):
                                                ts_str = self._format_timestamp(ts_raw)
                                                self._log(f"   [TS] Menggunakan timestamp dari Excel: {ts_str}")
                                            else:
                                                ts_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                                                self._log(f"   [TS] Kolom timestamp tidak ada/kosong — pakai waktu sekarang: {ts_str}")

                                            self._patch_timestamp(sheet, ts_str, row_before_submit)
                                    else:
                                        # Jika setelah 15 detik tetap tidak terdeteksi halaman konfirmasi
                                        raise SkipRespondentException("Gagal mendeteksi halaman konfirmasi pengiriman Google Form (timeout)")

                                    break
                                else: # Next button
                                    old_question = questions[0] if questions else None
                                    self._click_element(driver, nav_btn)
                                    halaman_ke += 1
                                    self._log(f"   -> Pindah ke Halaman {halaman_ke}")
                                    
                                    # Tunggu transisi halaman selesai (pertanyaan lama menghilang)
                                    if old_question:
                                        try:
                                            WebDriverWait(driver, 5).until(EC.staleness_of(old_question))
                                        except TimeoutException:
                                            pass
                                    
                                    # Tunggu halaman baru selesai terload (100%)
                                    try:
                                        WebDriverWait(driver, 10).until(
                                            lambda d: d.execute_script("return document.readyState") == "complete"
                                        )
                                    except TimeoutException:
                                        pass
                                        
                                    continue
                                    
                            except TimeoutException:
                                raise SkipRespondentException("Tidak ada tombol Navigasi (Next/Submit)")

                        if sukses_kirim:
                            break # Keluar dari loop RETRY, lanjut ke baris Excel berikutnya
                            
                    except SkipRespondentException as e:
                        # Skip langsung responden ini tanpa retry
                        alasan = str(e)
                        self._log(f"   [!!!] SKIP Resp. {index + 1}: {alasan}")
                        self._log(f"[ERROR_DATA]|{index + 1}|{nama_resp}|{alasan}")
                        sukses_kirim = False # Pastikan tidak dianggap sukses
                        break # Keluar dari loop RETRY, lanjut ke responden berikutnya
                        
                    except Exception as e:
                        # Tangkap error selenium (timeout, putus koneksi, dsb)
                        self._log(f"   [x] Error pada percobaan {attempt}: Terjadi kesalahan jaringan/halaman.")
                        time.sleep(2) # Jeda sejenak sebelum mencoba lagi
                
                else:
                    # else pada For-Loop akan dieksekusi jika loop berjalan sampai habis tanpa terkena 'break'
                    # (artinya semua percobaan gagal)
                    if not self._is_stopped():
                        self._log(f"   [CRITICAL] Gagal mengirim Resp. ke-{index + 1} setelah {MAX_RETRIES} kali percobaan. Melewati data ini.")
                        # nama_resp sudah di-set sebelum retry loop, tidak perlu re-assign
                        self._log(f"[ERROR_DATA]|{index + 1}|{nama_resp}|Gagal setelah {MAX_RETRIES} percobaan")

        finally:
            if driver:
                driver.quit()
            self._log("Selesai.")


