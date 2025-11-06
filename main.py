import os
import time
import pandas as pd
from datetime import datetime, timedelta
import zipfile
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import docx
import traceback
import re
import shutil
import subprocess
import requests

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# -------------------------------------------------------------------------
# GITHUB ACTIONS CONFIGURATION
# -------------------------------------------------------------------------
print("🚀 Initializing configuration for GitHub Actions environment...")
download_dir = os.path.join(os.getcwd(), "downloads_temp")
os.makedirs(download_dir, exist_ok=True)

options = webdriver.ChromeOptions()
user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
options.add_argument(f'user-agent={user_agent}')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
prefs = {"download.default_directory": download_dir, "download.prompt_for_download": False, "download.directory_upgrade": True}
options.add_experimental_option("prefs", prefs)
service = Service() 
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)
driver.set_page_load_timeout(60)
print("✅ WebDriver initialized successfully.")

# -------------------------------------------------------------------------
# HELPER FUNCTIONS
# -------------------------------------------------------------------------
PDF_PAGE_LIMIT = 10

def extract_text_from_pdf(file_path):
    text = ""
    try:
        doc = fitz.open(file_path)
        page_count = min(len(doc), PDF_PAGE_LIMIT)
        for i in range(page_count): text += doc[i].get_text("text") + "\n"
        doc.close()
    except Exception: return ""
    
    if len(text.strip()) < 50:
        print("    - Short text from PDF, attempting OCR fallback...")
        try:
            pages = convert_from_path(file_path, last_page=PDF_PAGE_LIMIT)
            ocr_text = ""
            for page_image in pages:
                ocr_text += pytesseract.image_to_string(page_image, lang="fra+ara+eng") + "\n"
            return ocr_text.strip()
        except Exception as e:
            print(f"    - ⚠️ OCR failed. Error: {e}")
            return ""
    return text.strip()

def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception: return ""

def extract_text_from_doc(file_path):
    try:
        process = subprocess.Popen(['antiword', file_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, _ = process.communicate()
        return stdout.decode('utf-8', errors='ignore')
    except Exception as e:
        print(f"    - ⚠️ Antiword failed for .doc file. Error: {e}")
        return ""
        
def extract_from_zip(file_path):
    try:
        extract_to = os.path.splitext(file_path)[0]
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(file_path, 'r') as zip_ref: zip_ref.extractall(extract_to)
        return extract_to
    except Exception as e:
        print(f"  - Failed to unzip {file_path}: {e}")
        return None

def clear_download_directory():
    for item_name in os.listdir(download_dir):
        item_path = os.path.join(download_dir, item_name)
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path): os.unlink(item_path)
            elif os.path.isdir(item_path): shutil.rmtree(item_path)
        except Exception as e: print(f"⚠️ Failed to delete {item_path}. Reason: {e}")

# *** CHANGE 1: A MUCH MORE ROBUST FUNCTION TO WAIT FOR DOWNLOADS ***
def wait_for_download_complete(timeout=90):
    """
    Waits for a download to complete. Ignores temporary .crdownload and .com.google.Chrome files.
    """
    seconds = 0
    while seconds < timeout:
        # Check for temporary files
        is_downloading = any(f.endswith('.crdownload') or f.startswith('.com.google.Chrome') for f in os.listdir(download_dir))
        
        if not is_downloading:
            # If no temp files, look for a completed file
            downloaded_files = [f for f in os.listdir(download_dir) if not (f.endswith('.crdownload') or f.startswith('.com.google.Chrome'))]
            if downloaded_files:
                return os.path.join(download_dir, downloaded_files[0]) # Return the first completed file

        time.sleep(1)
        seconds += 1
    return None # Return None if timeout is reached

# -------------------------------------------------------------------------
# MAIN SCRIPT LOGIC
# -------------------------------------------------------------------------
df = pd.DataFrame()

try:
    # --- PART 1: SCRAPE ALL TENDER METADATA INTO A DATAFRAME ---
    print("\n--- Starting Part 1: Scraping all tender metadata ---")
    URL1 = os.getenv("URL1")
    driver.get(URL1)
    print("🔑 Logging in...")
    wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_login"))).send_keys(os.getenv("USERNAME"))
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password").send_keys(os.getenv("PASSWORD"))
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton").click()
    print("✅ Login successful.")

    print("🔍 Navigating to search and applying filters...")
    time.sleep(3)
    URL2 = os.getenv("URL2")
    driver.get(URL2)
    date_input = wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")))
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
    date_input.clear(); date_input.send_keys(yesterday)
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche").click()
    print("✅ Search executed.")

    print("📊 Extracting data from results table...")
    wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")))
    Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")).select_by_value("500")
    time.sleep(3)
    
    rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')
    data = []
    for row in rows:
        try:
            data.append({
                "reference": row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text,
                "objet": row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", ""),
                "acheteur": row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", ""),
                "lieux_execution": row.find_element(By.XPATH, './/div[contains(@id,"panelBlocLieuxExec")]').text.replace("\n", ", "),
                "date_limite": row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text.replace("\n", " "),
                "download_page_url": row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")})
        except Exception: continue
    
    df = pd.DataFrame(data)
    excluded_words = ["construction", "installation", "recrutement", "travaux", "fourniture", "achat", "equipement", "maintenance", "works", "goods", "supply", "acquisition", "Recruitment", "nettoyage", "gardiennage"]
    if not df.empty:
        # df = df[~df['objet'].str.lower().str.contains('|'.join(excluded_words), na=False)]
        df_filtered = df[~df['objet'].str.lower().str.contains('|'.join(excluded_words), na=False)].reset_index(drop=True)

    print(f"✅ Found {len(df)} relevant tenders after filtering.")

    # --- PART 2: PROCESS EACH TENDER SEQUENTIALLY FROM THE DATAFRAME ---
    print("\n--- Starting Part 2: Processing each tender individually ---\n")
    all_processed_tenders = []
    
    for index, row in df_filtered.iterrows():
        print(f"--- Processing Tender {index + 1}/{len(df)} | Ref: {row['reference']} ---")
        try:
            clear_download_directory()
            
            # Step 2.1: ROBUST NAVIGATION AND DOWNLOAD
            driver.get(row['download_page_url'])
            
            dce_link = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dce_link); time.sleep(1)
            dce_link.click()

            checkbox = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox); time.sleep(1)
            checkbox.click()

            validate_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_validateButton")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", validate_btn); time.sleep(1)
            validate_btn.click()

            final_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", final_btn); time.sleep(1)
            final_btn.click()
            
            print("  - Download command sent. Waiting for file...")
            downloaded_file_path = wait_for_download_complete()
            
            # Step 2.2: EXTRACT TEXT FROM DOWNLOADED FILES
            merged_text = ""
            if not downloaded_file_path:
                print("  - ⚠️ Download failed or timed out. Skipping text extraction.")
                merged_text = "Error: Document download failed."
            else:
                print(f"  - ✅ Download complete: {os.path.basename(downloaded_file_path)}")
                file_paths_to_process = []
                if downloaded_file_path.lower().endswith(".zip"):
                    unzip_dir = extract_from_zip(downloaded_file_path)
                    if unzip_dir:
                        for r, _, files in os.walk(unzip_dir):
                            for f in files: file_paths_to_process.append(os.path.join(r, f))
                else:
                    file_paths_to_process.append(downloaded_file_path)
                
                print(f"  - Found {len(file_paths_to_process)} file(s) to process.")
                for file_path in file_paths_to_process:
                    filename = os.path.basename(file_path)
                    # *** CHANGE 2: IMPROVED LOGGING TO SHOW FILENAME ***
                    print(f"    -> Processing file: {filename}")

                    if "cps" in filename.lower():
                        print(f"       - Skipping file containing 'CPS'.")
                        continue
                    
                    ext = os.path.splitext(filename)[1].lower()
                    text = ""
                    if ext == ".pdf": text = extract_text_from_pdf(file_path)
                    elif ext == ".docx": text = extract_text_from_docx(file_path)
                    elif ext == ".doc": text = extract_text_from_doc(file_path)
                    
                    if text and text.strip():
                        print(f"       - Extracted {len(text)} characters.")
                        merged_text += f"\n\n--- Content from file: {filename} ---\n{text.strip()}"
                    else:
                        print("       - No text extracted.")
                
            # Step 2.3: PREPARE AND SEND DATA TO N8N
            tender_payload = row.to_dict()
            tender_payload['merged_text'] = merged_text.strip() if merged_text else "No relevant text could be extracted."
            tender_payload.pop('download_page_url', None)

            WEBHOOK_URL = os.getenv("N8N_WEBHOOK_URL")
            if WEBHOOK_URL:
                print(f"  - Sending data to n8n...")
                try:
                    response = requests.post(WEBHOOK_URL, json=tender_payload, timeout=30)
                    if response.status_code == 200: print(f"  - ✅ Successfully sent!")
                    else: print(f"  - ❌ Failed to send. Status: {response.status_code} (Check your n8n webhook URL and workflow status)")
                except Exception as e: print(f"  - ❌ Exception occurred while sending to n8n: {e}")
            
            all_processed_tenders.append(tender_payload)

        except Exception as e:
            print(f"  - ⚠️ An unexpected error occurred for this tender. Skipping. Error: {type(e).__name__}")
            traceback.print_exc()
            continue

finally:
    print("\n--- Finalizing Script ---")
    if all_processed_tenders:
        summary_df = pd.DataFrame(all_processed_tenders)
        output_csv_path = os.path.join(os.getcwd(), "tender_results_summary.csv")
        summary_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
        print(f"✅ Summary for {len(all_processed_tenders)} processed tenders saved to {output_csv_path}")
    else:
        print("ℹ️ No tenders were successfully processed.")
        
    print("Quitting WebDriver...")
    try: driver.quit()
    except Exception: pass

    if os.path.exists(download_dir):
        try:
            shutil.rmtree(download_dir)
            print("✅ Temporary download directory removed.")
        except Exception as e: print(f"⚠️ Could not remove download dir: {e}")

    print("🎉 Script finished successfully.")
