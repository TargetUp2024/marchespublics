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
import re  # Import the regular expressions module

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# Import the specific exception for timeouts
from selenium.common.exceptions import TimeoutException

# -----------------------------------------------------------------------------
# GITHUB ACTIONS CONFIGURATION
# -----------------------------------------------------------------------------
print("🚀 Initializing configuration for GitHub Actions environment...")
download_dir = os.path.join(os.getcwd(), "downloads", "Mp")
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
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)
service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)

# **NEW**: Set a specific timeout for page loads to avoid long hangs
driver.set_page_load_timeout(60) # Fail faster after 60 seconds

print("✅ WebDriver initialized successfully.")

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS (File Processing - All are correct)
# -----------------------------------------------------------------------------
PDF_PAGE_LIMIT = 10
def extract_text_from_pdf(file_path):
    text = ""
    try:
        doc = fitz.open(file_path)
        for i, page in enumerate(doc):
            if i >= PDF_PAGE_LIMIT: break
            text += page.get_text("text") + "\n"
        doc.close()
    except Exception: return ""
    if len(text.strip()) < 100:
        try:
            pages = convert_from_path(file_path, last_page=PDF_PAGE_LIMIT)
            ocr_text = ""
            for page_image in pages: ocr_text += pytesseract.image_to_string(page_image, lang="fra+ara") + "\n"
            return ocr_text.strip()
        except Exception: return ""
    return text.strip()

def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception: return ""

def extract_text_from_csv(file_path):
    try:
        df = pd.read_csv(file_path, on_bad_lines='skip', header=None, sep=None, engine='python')
        return df.to_string(index=False, header=False)
    except Exception as e:
        print(f"  - Error reading CSV {file_path}: {e}")
        return ""

def extract_text_from_xlsx(file_path):
    try:
        xls = pd.read_excel(file_path, sheet_name=None, header=None)
        full_text = []
        for sheet_name, df in xls.items():
            full_text.append(f"--- Sheet: {sheet_name} ---")
            full_text.append(df.to_string(index=False, header=False))
        return "\n".join(full_text)
    except Exception as e:
        print(f"  - Error reading Excel file {file_path}: {e}")
        return ""

def extract_from_zip(file_path):
    try:
        extract_to = os.path.splitext(file_path)[0]
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(file_path, 'r') as zip_ref: zip_ref.extractall(extract_to)
        print(f"  - Successfully unzipped {os.path.basename(file_path)}")
    except Exception as e: print(f"  - Failed to unzip {file_path}: {e}")

# -----------------------------------------------------------------------------
# MAIN SCRIPT LOGIC
# -----------------------------------------------------------------------------
try:
    # --- PART 1: WEB SCRAPING AND DOWNLOADING ---
    print("\n--- Starting Part 1: Web Scraping ---")
    
    # *** NEW: ADDED A RETRY LOOP FOR INITIAL CONNECTION ***
    MAX_RETRIES = 3
    for attempt in range(MAX_RETRIES):
        try:
            print(f"Attempting to connect to website (Attempt {attempt + 1}/{MAX_RETRIES})...")
            driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseHome")
            print("✅ Connection successful.")
            break # Exit the loop if successful
        except TimeoutException:
            print(f"⚠️ Page load timed out. Retrying in 10 seconds...")
            time.sleep(10)
            if attempt == MAX_RETRIES - 1: # If this was the last attempt
                print("❌ Failed to connect to the website after multiple retries. Aborting.")
                raise # Re-raise the last exception to stop the script
    
    print("🔑 Logging in...")
    wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_login"))).send_keys("TARGETUPCONSULTING")
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password").send_keys("pgwr00jPD@")
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton").click()
    print("✅ Login successful.")

    # (The rest of the script is unchanged as it was working correctly)
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
    time.sleep(2)
    rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')
    data = []
    for row in rows:
        try:
            ref_text = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
            objet = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", "")
            link = row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")
            ref_match = re.search(r'\d+', ref_text)
            if ref_match:
                ref_id = ref_match.group(0)
                data.append({"reference": ref_text, "ref_id": ref_id, "objet": objet, "download_page_url": link})
        except Exception: continue
    df = pd.DataFrame(data)
    excluded_words = ["construction", "installation", "recrutement", "travaux", "fourniture", "achat", "equipement", "maintenance", "works", "goods", "supply", "acquisition", "Recruitment", "nettoyage", "gardiennage"]
    df = df[~df['objet'].str.lower().str.contains('|'.join(excluded_words), na=False)]
    print(f"✅ Found {len(df)} relevant tenders after filtering.")

    links_to_process = df['download_page_url'].tolist()[:5]
    print(f"\n📥 Starting download loop for the first {len(links_to_process)} links...")
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
            print(f"✅ Download initiated. Waiting 15 seconds for file to save...")
            time.sleep(15)
        except Exception:
            error_filename = f"error_page_{i+1}.png"
            print(f"⚠️ An error occurred for link {link}. Saving screenshot to {error_filename}")
            driver.save_screenshot(error_filename)
            with open(f"error_page_{i+1}.html", "w", encoding="utf-8") as f: f.write(driver.page_source)
            traceback.print_exc()
            continue
    print("\n🎯 Download loop finished.")

    print("\n--- Starting Part 2: File Processing ---")
    print("Step 2.1: Unzipping all downloaded .zip files...")
    for root, _, files in os.walk(download_dir):
        for f in files:
            if f.lower().endswith(".zip"): extract_from_zip(os.path.join(root, f))
    
    print("\nStep 2.2: Extracting text from all files...")
    tender_results = []
    for item_name in os.listdir(download_dir):
        item_path = os.path.join(download_dir, item_name)
        if not os.path.isdir(item_path): continue
        print(f"\nProcessing folder: {item_name}")
        merged_text = ""
        ref_match = re.search(r'\d+', item_name)
        if not ref_match:
            print(f"  - Warning: Could not find a reference number in folder name '{item_name}'. Skipping.")
            continue
        ref_id = ref_match.group(0)
        for root, _, files in os.walk(item_path):
            for f in files:
                if 'cps' in f.lower(): continue
                file_path = os.path.join(root, f)
                ext = os.path.splitext(f)[1].lower()
                text = ""
                if ext == ".pdf": text = extract_text_from_pdf(file_path)
                elif ext == ".docx": text = extract_text_from_docx(file_path)
                elif ext == ".csv": text = extract_text_from_csv(file_path)
                elif ext in [".xls", ".xlsx"]: text = extract_text_from_xlsx(file_path)
                if text and text.strip(): merged_text += f"\n\n{'='*20}\n--- Content from file: {f} ---\n{'='*20}\n{text.strip()}"
        if merged_text.strip(): tender_results.append({"ref_id": ref_id, "merged_text": merged_text.strip()})
    
    df1 = pd.DataFrame(tender_results)

    print("\n--- Starting Part 3: Merging data and saving to CSV ---")
    if not df.empty and not df1.empty:
        merged_df = pd.merge(df, df1, on="ref_id", how="inner")
        if 'ref_id' in merged_df.columns: merged_df = merged_df.drop(columns=['ref_id'])
        output_csv_path = "tender_results.csv"
        merged_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
        print(f"✅ Data successfully merged and saved to {output_csv_path}")
    else:
        print("⚠️ Scraped data or processed file data is empty. Cannot merge. Saving initial scrape results only.")
        df.to_csv("tender_results.csv", index=False, encoding='utf-8-sig')

finally:
    print("\nQuitting WebDriver...")
    driver.quit()
    print("🎉 Script finished successfully.")
