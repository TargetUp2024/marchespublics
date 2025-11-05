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
# Safety: ensure merged_df exists even if try block fails early
merged_df = pd.DataFrame()

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
    if len(text.strip()) < 50:
        try:
            pages = convert_from_path(file_path, last_page=PDF_PAGE_LIMIT)
            ocr_text = ""
            for page_image in pages:
                ocr_text += pytesseract.image_to_string(page_image, lang="fra+ara+eng") + "\n"
            return ocr_text.strip()
        except Exception:
            return ""
    return text.strip()

def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception:
        return ""

def extract_text_from_doc(file_path):
    try:
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
    try:
        extract_to = os.path.splitext(file_path)[0]
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
        print(f"  - Successfully unzipped {os.path.basename(file_path)}")
    except Exception as e:
        print(f"  - Failed to unzip {file_path}: {e}")

# -------------------------------------------------------------------------
# MAIN SCRIPT LOGIC
try:
    # --- PART 1: WEB SCRAPING ---
    print("\n--- Starting Part 1: Web Scraping ---")
    MAX_RETRIES = 3
    for attempt in range(MAX_RETRIES):
        try:
            print(f"Attempting to connect to website (Attempt {attempt + 1}/{MAX_RETRIES})...")
            driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseHome")
            break
        except TimeoutException:
            if attempt == MAX_RETRIES - 1:
                raise
            print("⚠️ Page load timed out. Retrying in 10 seconds...")
            time.sleep(10)

    print("🔑 Logging in...")
    wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_login"))).send_keys("TARGETUPCONSULTING")
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password").send_keys("pgwr00jPD@")
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton").click()
    print("✅ Login successful.")

    print("🔍 Navigating to advanced search and setting filters...")
    time.sleep(5)
    driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons")
    date_input = wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")))
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
    date_input.clear()
    date_input.send_keys(yesterday)
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
            ref_text = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
            objet = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", "")
            buyer = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", "")
            lieux = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocLieuxExec")]').text.replace("\n", ", ")
            deadline = row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text.replace("\n", " ")
            link = row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")
            ref_match = re.search(r'\d+', ref_text)
            if ref_match:
                ref_id = ref_match.group(0)
                data.append({
                    "reference": ref_text,
                    "ref_id": ref_id,
                    "objet": objet,
                    "acheteur": buyer,
                    "lieux_execution": lieux,
                    "date_limite": deadline,
                    "download_page_url": link
                })
        except Exception:
            continue

    df = pd.DataFrame(data)
    excluded_words = [
        "construction", "installation", "recrutement", "travaux",
        "fourniture", "achat", "equipement", "maintenance",
        "works", "goods", "supply", "acquisition", "Recruitment",
        "nettoyage", "gardiennage"
    ]
    if not df.empty:
        df = df[~df['objet'].str.lower().str.contains('|'.join(excluded_words), na=False)]
    print(f"✅ Found {len(df)} relevant tenders after filtering.")

    # --- PART 2: DOWNLOADING ---
    links_to_process = df['download_page_url'].tolist() if not df.empty else []
    ## links_to_process = links_to_process[:5]
    print(f"\n📥 Starting download loop for {len(links_to_process)} tenders...")
    for i, link in enumerate(links_to_process):
        print(f"\n--- Processing link {i+1}/{len(links_to_process)} ---")
        try:
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

            print("✅ Download initiated. Waiting 15 seconds...")
            time.sleep(15)
        except Exception:
            error_filename = f"error_page_{i+1}.png"
            print(f"⚠️ An error occurred during download. Saving screenshot to {error_filename}")
            driver.save_screenshot(error_filename)
            with open(f"error_page_{i+1}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            traceback.print_exc()
            continue
    print("\n🎯 Download loop finished.")

    # --- PART 3: FILE PROCESSING ---
    print("\n--- Starting Part 3: File Processing ---")
    print("Step 3.1: Unzipping all downloaded .zip files...")
    for root, _, files in os.walk(download_dir):
        for f in files:
            if f.lower().endswith(".zip"):
                extract_from_zip(os.path.join(root, f))

    print("\nStep 3.2: Extracting text from all files...")
    tender_results = []
    for item_name in os.listdir(download_dir):
        item_path = os.path.join(download_dir, item_name)
        if not os.path.isdir(item_path):
            continue
        print(f"\n📂 Tender Folder: {item_name}")
        file_counter = 1
        merged_text = ""
        ref_match = re.search(r'\d+', item_name)
        if not ref_match:
            continue
        ref_id = ref_match.group(0)

        for root, _, files in os.walk(item_path):
            for f in files:
                print(f"   ➜ Extracting file {file_counter}: {f}")
                file_counter += 1
                file_path = os.path.join(root, f)
                ext = os.path.splitext(f)[1].lower()
                text = ""
                if ext == ".pdf":
                    text = extract_text_from_pdf(file_path)
                elif ext == ".docx":
                    text = extract_text_from_docx(file_path)
                elif ext == ".doc":
                    text = extract_text_from_doc(file_path)
                elif ext == ".csv":
                    text = extract_text_from_csv(file_path)
                elif ext in [".xls", ".xlsx"]:
                    text = extract_text_from_xlsx(file_path)

                if text and text.strip():
                    merged_text += f"\n\n{'='*20}\n--- Content from file: {f} ---\n{'='*20}\n{text.strip()}"

        if merged_text.strip():
            tender_results.append({"ref_id": ref_id, "merged_text": merged_text.strip()})

    # --- Build df1 & collapse multiple text rows per ref_id ---
    df1 = pd.DataFrame(tender_results)
    if not df1.empty:
        df1['ref_id'] = df1['ref_id'].astype(str)
        df1 = df1.groupby('ref_id', as_index=False).agg({
            'merged_text': lambda texts: "\n\n".join(dict.fromkeys([t for t in texts if t and str(t).strip()]))
        })

    # --- PART 4: MERGE AND SAVE TO CSV ---
    print("\n--- Starting Part 4: Merging data and saving to CSV ---")

    # Ensure scrape results also unique per tender
    if not df.empty:
        df['ref_id'] = df['ref_id'].astype(str)
        df = df.drop_duplicates(subset=['ref_id'])

    workspace_path = os.getcwd()
    output_csv_path = os.path.join(workspace_path, "tender_results.csv")

    if not df.empty and not df1.empty:
        merged_df = pd.merge(df, df1, on="ref_id", how="inner")
        if 'ref_id' in merged_df.columns:
            merged_df = merged_df.drop(columns=['ref_id'])
        if 'download_page_url' in merged_df.columns:
            merged_df = merged_df.drop(columns=['download_page_url'])
        merged_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
        print(f"✅ Data for {len(merged_df)} tenders successfully merged and saved to {output_csv_path}")

    elif not df.empty:
        # Keep scraped tenders (no attachments)
        merged_df = df.copy()
        merged_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
        print(f"⚠️ No processed attachments. Saved {len(merged_df)} scraped tenders to {output_csv_path}")

    else:
        # No scraped data at all - create empty CSV
        merged_df = pd.DataFrame()
        merged_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
        print("❌ No data was scraped. Created an empty CSV.")

finally:
    # Safe finally - merged_df is guaranteed to exist (could be empty)
    print("\nQuitting WebDriver and cleaning up...")
    try:
        json_data = merged_df.to_dict(orient="records")
    except Exception:
        json_data = []
    print("Ready for N8N !!!!!!!")
    time.sleep(1)

    webhook_url = "https://targetup.app.n8n.cloud/webhook/78e3201b-36a3-4341-a067-e74f0693be6d"
    try:
        response = requests.post(webhook_url, json=json_data, timeout=20)
        if response.status_code == 200:
            print("✅ Data sent successfully!")
        else:
            print(f"❌ Failed to send data. Status code: {response.status_code}")
    except Exception as e:
        print(f"❌ Exception while sending webhook: {e}")

    try:
        driver.quit()
    except Exception:
        pass

    if os.path.exists(download_dir):
        try:
            shutil.rmtree(download_dir)
            print("✅ Temporary download directory removed.")
        except Exception as e:
            print(f"⚠️ Could not remove download dir: {e}")

    print("🎉 Script finished successfully.")
