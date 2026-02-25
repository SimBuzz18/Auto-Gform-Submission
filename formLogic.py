import pandas as pd
import time
import re
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager


class SkipRespondentException(Exception):
    """Exception raised to skip current respondent and move to next."""
    pass


def worker_launcher(df_chunk, link_gform, queue_log, stop_event, worker_id, headless=True):
    """Entry point for multiprocessing worker."""
    form_logic = logic(callback_log=queue_log, stop_event=stop_event, worker_id=worker_id)
    form_logic.run_process_chunk(df_chunk, link_gform, headless=headless)
    

class logic:
    def __init__(self, callback_log=None, stop_event=None, worker_id=None):
        """
        callback_log: can be a function (threading) or a multiprocessing.Queue
        stop_event: multiprocessing.Event or None
        worker_id: identifier for logging
        """
        self.callback_log = callback_log
        self.stop_event = stop_event
        self.worker_id = worker_id
        self.separator_checkbox = ','

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

    def run_process_chunk(self, df_chunk, link_gform, headless=False):
        """Runs the automation for a specific chunk of data."""
        driver = None
        try:
            # Normalize column names
            col_map = {self.normalize_text(col): col for col in df_chunk.columns}
            
            # Setup Browser
            self._log("Membuka Browser Chrome...")
            options = webdriver.ChromeOptions()
            if headless:
                options.add_argument("--headless")
            
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

            # ... (kode setup browser sebelumnya tetap sama) ...

            MAX_RETRIES = 3 # Maksimal percobaan per responden

            for index, row in df_chunk.iterrows():
                if self._is_stopped():
                    self._log("[STOP] Proses dihentikan.")
                    break

                for attempt in range(1, MAX_RETRIES + 1):
                    if self._is_stopped(): break
                    
                    if attempt == 1:
                        self._log(f"Resp. ke-{index + 1}...")
                    else:
                        self._log(f"   [!] Retrying Resp. ke-{index + 1} (Percobaan {attempt}/{MAX_RETRIES})...")

                    try:
                        nama_resp = row.get('Nama Lengkap', row.get('Nama', 'Unknown'))
                        driver.get(link_gform)
                        wait = WebDriverWait(driver, 10) 
                        
                        halaman_ke = 1
                        sukses_kirim = False # Penanda apakah berhasil sampai tombol Kirim
                        
                        while True:
                            if self._is_stopped(): break
                            
                            # Update: Use dynamic wait for questions instead of fixed sleep
                            try:
                                wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[@role='listitem']")))
                            except TimeoutException:
                                self._log("   [!] Timeout menunggu pertanyaan muncul.")
                                break

                            questions = driver.find_elements(By.XPATH, "//div[@role='listitem']")
                            
                            # (Loop pengisian pertanyaan tetap sama seperti sebelumnya)
                            for q_elem in questions:
                                try:
                                    try:
                                        heading_elem = q_elem.find_element(By.XPATH, ".//div[@role='heading']/span")
                                        q_text = heading_elem.text
                                    except:
                                        continue 

                                    normalized_q_text = self.normalize_text(q_text)
                                    target_col = None
                                    
                                    if normalized_q_text in col_map:
                                        target_col = col_map[normalized_q_text]
                                    
                                    if not target_col:
                                        raise SkipRespondentException(f"Pertanyaan '{q_text}' tidak ditemukan di Excel")
                                        
                                    answer = row[target_col]
                                    if pd.isna(answer) or str(answer).lower() == 'nan' or str(answer).strip() == '':
                                        continue 

                                    # A. Radio Button
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
                                                    driver.execute_script("arguments[0].click();", radio)
                                                    time.sleep(0.3)
                                                found_radio = True
                                                break
                                        
                                        if not found_radio:
                                            raise SkipRespondentException(f"Pilihan '{answer}' tidak ada untuk pertanyaan '{q_text}'")
                                        continue

                                    # B. Checkbox
                                    checkboxes = q_elem.find_elements(By.XPATH, ".//div[@role='checkbox']")
                                    if checkboxes:
                                        answers_list = [self.normalize_text(x) for x in str(answer).split(self.separator_checkbox)]
                                        found_count = 0 
                                        
                                        for cb in checkboxes:
                                            data_val = cb.get_attribute("data-value")
                                            aria_label = cb.get_attribute("aria-label")
                                            
                                            val_norm = ""
                                            if data_val: val_norm = self.normalize_text(data_val)
                                            elif aria_label: val_norm = self.normalize_text(aria_label)
                                            
                                            if val_norm in answers_list:
                                                if cb.get_attribute("aria-checked") == "false":
                                                    driver.execute_script("arguments[0].click();", cb)
                                                    time.sleep(0.3)
                                                found_count += 1
                                        
                                        if found_count < len(answers_list):
                                            raise SkipRespondentException(f"Opsi checkbox '{answer}' tidak lengkap di Form")
                                        continue

                                    # C. Text Input
                                    text_inputs = q_elem.find_elements(By.XPATH, ".//input[@type='text'] | .//textarea")
                                    if text_inputs:
                                        inp = text_inputs[0]
                                        inp.clear()
                                        inp.send_keys(str(answer))
                                        continue
                                        
                                except SkipRespondentException:
                                    raise # Re-raise to skip respondent
                                except Exception as e:
                                    # Perlakukan error interaksi/selenium sebagai pemicu skip agar tidak menekan 'Kirim' setengah jadi
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
                                    driver.execute_script("arguments[0].click();", nav_btn)
                                    self._log(f"[v] Resp. {index + 1} TERKIRIM.")
                                    time.sleep(2)
                                    sukses_kirim = True
                                    break
                                else: # Next button
                                    driver.execute_script("arguments[0].scrollIntoView();", nav_btn)
                                    time.sleep(0.3)
                                    driver.execute_script("arguments[0].click();", nav_btn)
                                    halaman_ke += 1
                                    self._log(f"   -> Pindah ke Halaman {halaman_ke}")
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
                        # Kirim data error spesifik untuk rekap di UI
                        nama_resp = row.get('Nama Lengkap', row.get('Nama', 'Unknown'))
                        # Format: [ERROR_DATA]|Baris|Nama|Alasan
                        self._log(f"[ERROR_DATA]|{index + 1}|{nama_resp}|Gagal setelah {MAX_RETRIES} percobaan")

        finally:
            if driver:
                driver.quit()
            self._log("Selesai.")


