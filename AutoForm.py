import pandas as pd
import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager

class AutoFormLogic:
    def __init__(self, callback_log, callback_is_stopped):
        self.log = callback_log
        self.is_stopped = callback_is_stopped
        self.separator_checkbox = ','

    def clean_data(self, val):
        val = str(val).strip()
        if val.endswith('.0'):
            return val[:-2]
        return val

    def normalize_text(self, text):
        if not isinstance(text, str): return ""
        return re.sub(r'\s+', ' ', text.lower().strip())

    def run_process(self, file_excel, link_gform):
        driver = None
        try:
            # 1. Baca Excel
            try:
                self.log("Membaca file Excel...")
                df = pd.read_excel(file_excel)
                df = df.map(self.clean_data)
                self.log(f"Berhasil membaca {len(df)} baris data.")
                col_map = {self.normalize_text(col): col for col in df.columns}
            except Exception as e:
                self.log(f"Error membaca Excel: {e}")
                return

            # 2. Buka Browser
            self.log("Membuka Browser Chrome...")
            options = webdriver.ChromeOptions()
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

            # 3. Loop Responden
            for index, row in df.iterrows():
                if self.is_stopped():
                    self.log("[STOP] Proses dihentikan oleh pengguna.")
                    break

                self.log(f"\n=== Mengisi Responden ke-{index + 1} ===")
                try:
                    driver.get(link_gform)
                    wait = WebDriverWait(driver, 10) 
                    
                    halaman_ke = 1
                    while True:
                        if self.is_stopped(): break
                        
                        time.sleep(2)
                        self.log(f"   > Halaman {halaman_ke}...")
                        
                        questions = driver.find_elements(By.XPATH, "//div[@role='listitem']")
                        
                        for q_elem in questions:
                            try:
                                # Cari Teks Pertanyaan
                                try:
                                    heading_elem = q_elem.find_element(By.XPATH, ".//div[@role='heading']/span")
                                    q_text = heading_elem.text
                                except:
                                    continue 

                                normalized_q_text = self.normalize_text(q_text)
                                target_col = None
                                
                                # Match Column Logic
                                if normalized_q_text in col_map:
                                    target_col = col_map[normalized_q_text]
                                else:
                                    for norm_col, real_col in col_map.items():
                                        if norm_col in normalized_q_text or normalized_q_text in norm_col:
                                            target_col = real_col
                                            break
                                
                                if not target_col:
                                    # CRITICAL: Pertanyaan di Form tidak ada di Excel -> STOP
                                    self.log(f"[CRITICAL] Pertanyaan '{q_text}' TIDAK DITEMUKAN di Excel!")
                                    self.log("Menghentikan proses otomatis...")
                                    return # Stop total
                                    
                                answer = row[target_col]
                                if pd.isna(answer) or answer == '' or answer.lower() == 'nan':
                                    continue 

                                # --- DETEKSI INPUT & STRICT VALIDATION ---
                                
                                # A. Radio Button
                                radios = q_elem.find_elements(By.XPATH, ".//div[@role='radio']")
                                if radios:
                                    found_radio = False
                                    for radio in radios:
                                        data_val = radio.get_attribute("data-value")
                                        aria_label = radio.get_attribute("aria-label")
                                        
                                        is_match = False
                                        if data_val and self.normalize_text(data_val) == self.normalize_text(answer): is_match = True
                                        elif aria_label and self.normalize_text(aria_label) == self.normalize_text(answer): is_match = True
                                        
                                        if is_match:
                                            if radio.get_attribute("aria-checked") == "false":
                                                driver.execute_script("arguments[0].click();", radio)
                                                time.sleep(0.3)
                                            found_radio = True
                                            break
                                    
                                    if not found_radio:
                                        # Strict Check Radio
                                        self.log(f"[CRITICAL] Pilihan '{answer}' tidak ada untuk pertanyaan '{q_text}' (Radio)")
                                        self.log("Menghentikan proses...")
                                        return
                                    else:
                                        # Info Success (Behind scenes)
                                        # self.log(f"      [OK] Radio valid: {answer}")
                                        pass
                                    continue

                                # B. Checkbox
                                checkboxes = q_elem.find_elements(By.XPATH, ".//div[@role='checkbox']")
                                if checkboxes:
                                    answers_list = [self.normalize_text(x) for x in answer.split(self.separator_checkbox)]
                                    # Check coverage: Setiap jawaban di excel harus ada di form
                                    found_count = 0 
                                    
                                    # Reset dulu (optional, but good for idempotency) - skip for now
                                    
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
                                        # Artinya ada jawaban di excel yang tidak ketemu di checkbox form
                                        self.log(f"[CRITICAL] Salah satu opsi checkbox '{answer}' tidak ditemukan di Form!")
                                        return
                                    continue

                                # C. Text Input
                                text_inputs = q_elem.find_elements(By.XPATH, ".//input[@type='text'] | .//textarea")
                                if text_inputs:
                                    inp = text_inputs[0]
                                    if inp.get_attribute("value") != answer:
                                        inp.clear()
                                        inp.send_keys(answer)
                                    continue
                                    
                            except Exception as e:
                                self.log(f"Error handling question: {e}")
                                # continue

                        # --- NAVIGASI ---
                        if self.is_stopped(): break

                        # Submit / Next Logic (Same as before)
                        try:
                            submit_btn = driver.find_element(By.XPATH, "//div[@role='button']//span[text()='Kirim' or text()='Submit']")
                            driver.execute_script("arguments[0].scrollIntoView();", submit_btn)
                            time.sleep(0.5)
                            submit_btn.click()
                            self.log("   [v] DATA TERKIRIM.")
                            time.sleep(3)
                            break 
                        except NoSuchElementException:
                            pass
                        
                        try:
                            next_btn = driver.find_element(By.XPATH, "//div[@role='button']//span[text()='Berikutnya' or text()='Next']")
                            driver.execute_script("arguments[0].scrollIntoView();", next_btn)
                            time.sleep(0.5)
                            next_btn.click()
                            self.log("   -> Halaman berikutnya...")
                            halaman_ke += 1
                            time.sleep(2)
                            continue
                        except NoSuchElementException:
                            self.log("   [?] Tidak ada tombol Next/Submit.")
                            break

                except Exception as e:
                    self.log(f"Error Responden Loop: {e}")
        
        finally:
            self.log("=== PROSES SELESAI / BERHENTI ===")
