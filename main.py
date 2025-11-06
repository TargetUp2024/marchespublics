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
from selenium.common.exceptions import TimeoutException, NoSuchElementException

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
# HELPER FUNCTIONS - TEXT EXTRACTION
# -------------------------------------------------------------------------
PDF_PAGE_LIMIT = 10

def extract_text_from_pdf(file_path):
    text = ""
    try:
        doc = fitz.open(file_path)
        page_count = min(len(doc), PDF_PAGE_LIMIT)
        for i in range(page_count):
            text += doc[i].get_text("text") + "\n"
        doc.close()
    except Exception:
        return ""
    if len(text.strip()) < 50: # If text extraction is poor, try OCR
        try:
            pages = convert_from_path(file_path, last_page=PDF_PAGE_LIMIT)
            ocr_text = ""
            for page_image in pages:
                ocr_text += pytesseract.image_to_string(page_image, lang="fra+ara+eng") + "\n"
            return ocr_text.strip()
        except Exception:
            return "" # Return empty if OCR also fails
    return text.strip()

def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception:
        return ""

def extract_text_from_doc(file_path):
    try:
        # Use antiword for .doc files
        process = subprocess.Popen(['antiword', file_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, _ = process.communicate()
        return stdout.decode('utf-8', errors='ignore')
    except Exception as e:
        print(f"  - Error reading .doc file with antiword: {e}")
        return ""

def extract_text_from_csv(file_path):
    try:
        df = pd.read_csv(file_path, on_bad_lines='skip', header=None, sep=None, engine='python')
        return df.to_string(index=False, header=False)
    except Exception:
        return ""

def extract_text_from_xlsx(file_path):
    try:
        xls = pd.read_excel(file_path, sheet_name=None, header=None)
        full_text = []
        for sheet_name, df in xls.items():
            full_text.append(f"--- Sheet: {sheet_name} ---")
            full_text.append(df.to_string(index=False, header=False))
        return "\n".join(full_text)
    except Exception:
        return ""

def extract_from_zip(file_path):
    """Extracts a zip file and returns the directory path."""
    try:
        extract_to = os.path.splitext(file_path)[0]
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
        print(f"  - Successfully unzipped to {extract_to}")
        return extract_to
    except Exception as e:
        print(f"  - Failed to unzip {file_path}: {e}")
        return None

def clear_download_directory():
    """Removes all files and folders in the download directory."""
    for item_name in os.listdir(download_dir):
        item_path = os.path.join(download_dir, item_name)
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path):
                os.unlink(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
        except Exception as e:
            print(f"⚠️ Failed to delete {item_path}. Reason: {e}")

def wait_for_download_complete(timeout=60):
    """Waits for a download to complete by checking for .crdownload files."""
    seconds = 0
    while seconds < timeout:
        crdownload_exists = any(f.endswith('.crdownload') for f in os.listdir(download_dir))
        if not crdownload_exists:
            # Find the first non-crdownload file and return its path
            for f in os.listdir(download_dir):
                if not f.endswith('.crdownload'):
                    return os.path.join(download_dir, f)
            # If directory is empty, the download might not have started
        time.sleep(1)
        seconds += 1
    # If timeout is reached, return None
    return None

# -------------------------------------------------------------------------
# MAIN SCRIPT LOGIC
# -------------------------------------------------------------------------
all_tender_data = [] # To store all processed data for final CSV logging

try:
    # --- PART 1: WEB SCRAPING SETUP ---
    print("\n--- Starting Part 1: Web Scraping Setup ---")
    MAX_RETRIES = 3
    for attempt in range(MAX_RETRIES):
        try:
            print(f"Attempting to connect to website (Attempt {attempt + 1}/{MAX_RETRIES})...")
            URL1 = os.getenv("URL1")
            driver.get(URL1)
            break # Exit loop if successful
        except TimeoutException:
            if attempt == MAX_RETRIES - 1:
                raise # Raise the final timeout error
            print("⚠️ Page load timed out. Retrying in 10 seconds...")
            time.sleep(10)

    print("🔑 Logging in...")
    USERNAME = os.getenv("USERNAME")
    wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_login"))).send_keys(USERNAME)
    PASSWORD = os.getenv("PASSWORD")
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password").send_keys(PASSWORD)
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton").click()
    print("✅ Login successful.")

    print("🔍 Navigating to advanced search and setting filters...")
    time.sleep(5)
    URL2 = os.getenv("URL2")
    driver.get(URL2)
    date_input = wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")))
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
    date_input.clear()
    date_input.send_keys(yesterday)
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche").click()
    print("✅ Search executed.")

    print("📊 Preparing to process results one by one...")
    wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")))
    Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")).select_by_value("500")
    time.sleep(3) # Wait for the page to reload with 500 results
    rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')

    excluded_words = [
        "construction", "installation", "recrutement", "travaux",
        "fourniture", "achat", "equipement", "maintenance",
        "works", "goods", "supply", "acquisition", "Recruitment",
        "nettoyage", "gardiennage"
    ]

    print(f"Found {len(rows)} potential tenders. Starting individual processing loop...")

    # --- PART 2: INDIVIDUAL TENDER PROCESSING LOOP ---
    for i, row in enumerate(rows):
        tender_data = {}
        try:
            print(f"\n--- Processing Tender {i+1}/{len(rows)} ---")

            # Step 2.1: Scrape basic data from the row
            objet = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", "")

            # Filter out excluded words
            if any(word in objet.lower() for word in excluded_words):
                print(f"Skipping tender due to excluded keyword in 'objet': {objet[:50]}...")
                continue

            ref_text = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
            buyer = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", "")
            lieux = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocLieuxExec")]').text.replace("\n", ", ")
            deadline = row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text.replace("\n", " ")
            link = row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")

            tender_data = {
                "reference": ref_text,
                "objet": objet,
                "acheteur": buyer,
                "lieux_execution": lieux,
                "date_limite": deadline
            }
            print(f"Scraped data for Ref: {ref_text}")

            # Step 2.2: Download documents
            print("  - Navigating to download page...")
            clear_download_directory() # Ensure directory is clean before download
            driver.get(link)
            download_link = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", download_link)
            time.sleep(0.5)
            download_link.click()

            checkbox = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
            time.sleep(0.5)
            checkbox.click()

            valider_button = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_validateButton")))
            valider_button.click()

            final_download_button = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload")))
            final_download_button.click()
            print("  - Download initiated. Waiting for completion...")

            downloaded_file_path = wait_for_download_complete()
            if not downloaded_file_path:
                print("  - ⚠️ Download failed or timed out. Skipping file processing.")
                tender_data["merged_text"] = "Error: Document download failed."
            else:
                print(f"  - ✅ Download complete: {os.path.basename(downloaded_file_path)}")

                # Step 2.3: Process downloaded files
                merged_text = ""
                file_paths_to_process = []
                
                if downloaded_file_path.lower().endswith(".zip"):
                    unzip_dir = extract_from_zip(downloaded_file_path)
                    if unzip_dir:
                        for root, _, files in os.walk(unzip_dir):
                            for f in files:
                                file_paths_to_process.append(os.path.join(root, f))
                else:
                    file_paths_to_process.append(downloaded_file_path)

                print(f"  - Found {len(file_paths_to_process)} file(s) to process.")
                for file_path in file_paths_to_process:
                    filename = os.path.basename(file_path)
                    print(f"    ➜ Extracting text from: {filename}")
                    ext = os.path.splitext(filename)[1].lower()
                    text = ""
                    if ext == ".pdf": text = extract_text_from_pdf(file_path)
                    elif ext == ".docx": text = extract_text_from_docx(file_path)
                    elif ext == ".doc": text = extract_text_from_doc(file_path)
                    elif ext == ".csv": text = extract_text_from_csv(file_path)
                    elif ext in [".xls", ".xlsx"]: text = extract_text_from_xlsx(file_path)

                    if text and text.strip():
                        merged_text += f"\n\n{'='*20}\n--- Content from file: {filename} ---\n{'='*20}\n{text.strip()}"
                
                tender_data["merged_text"] = merged_text.strip() if merged_text else "No text could be extracted from documents."

            # Step 2.4: Send data to n8n Webhook
            WEBHOOK_URL = os.getenv("N8N_WEBHOOK_URL")
            if WEBHOOK_URL:
                print(f"  - Sending data for Ref: {ref_text} to n8n...")
                try:
                    response = requests.post(WEBHOOK_URL, json=tender_data, timeout=30)
                    if response.status_code == 200:
                        print(f"  - ✅ Successfully sent to n8n!")
                    else:
                        print(f"  - ❌ Failed to send to n8n. Status: {response.status_code} | Response: {response.text}")
                except Exception as e:
                    print(f"  - ❌ An exception occurred while sending to n8n: {e}")
            else:
                print("  - ⚠️ N8N_WEBHOOK_URL not set. Skipping sending.")
            
            # Append to the list for final CSV logging
            all_tender_data.append(tender_data)

        except (NoSuchElementException, TimeoutException) as e:
            print(f"⚠️ A critical error occurred for tender {i+1} (Ref: {tender_data.get('reference', 'N/A')}). Skipping. Error: {e}")
            traceback.print_exc()
            # Navigate back to search results to prevent being stuck on an error page
            driver.get(URL2) 
            # Re-apply search to get back to the list
            wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche"))).click()
            wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")))
            Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")).select_by_value("500")
            time.sleep(3)
            # Re-fetch rows to continue from the next item
            rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')
            continue
        except Exception as e:
            print(f"An unexpected error occurred processing tender {i+1}: {e}")
            traceback.print_exc()
            continue
        finally:
            # Clean up downloads for the next iteration
            clear_download_directory()

finally:
    print("\n--- Finalizing Script ---")
    
    # --- Create a final summary CSV file ---
    if all_tender_data:
        final_df = pd.DataFrame(all_tender_data)
        output_csv_path = os.path.join(os.getcwd(), "tender_results_summary.csv")
        final_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
        print(f"✅ Summary for {len(final_df)} processed tenders saved to {output_csv_path}")
    else:
        print("ℹ️ No tenders were successfully processed to create a summary file.")
        
    print("Quitting WebDriver...")
    try:
        driver.quit()
    except Exception:
        pass

    if os.path.exists(download_dir):
        try:
            # Final cleanup of the main download directory
            shutil.rmtree(download_dir)
            print("✅ Temporary download directory removed.")
        except Exception as e:
            print(f"⚠️ Could not remove download dir on final cleanup: {e}")

    print("🎉 Script finished successfully.")
