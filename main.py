import os
import time
import random
import pandas as pd
from datetime import datetime, timedelta
import zipfile
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import docx
import traceback

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# -----------------------------------------------------------------------------
# GITHUB ACTIONS CONFIGURATION
# -----------------------------------------------------------------------------
print("🚀 Initializing configuration for GitHub Actions environment...")
download_dir = os.path.join(os.getcwd(), "downloads", "Mp")
os.makedirs(download_dir, exist_ok=True)

options = webdriver.ChromeOptions()

# --- MAKE THE BOT LOOK MORE HUMAN TO AVOID DETECTION ---
user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
options.add_argument(f'user-agent={user_agent}')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# --- STANDARD ARGUMENTS FOR HEADLESS LINUX RUNNER ---
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")

# --- SET DOWNLOAD PREFERENCES ---
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)

# --- INITIALIZE WEB DRIVER ---
service = Service()
driver = webdriver.Chrome(service=service, options=options)
# Use a longer wait time for potentially slow runners
wait = WebDriverWait(driver, 20)
print("✅ WebDriver initialized successfully.")

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS (File Processing)
# -----------------------------------------------------------------------------
# NOTE: In a GitHub Actions runner, Tesseract and Poppler are installed system-wide.
# We do not need to specify paths to their executables.

PDF_PAGE_LIMIT = 10

def extract_text_from_pdf(file_path):
    text = ""
    try:
        doc = fitz.open(file_path)
        for i, page in enumerate(doc):
            if i >= PDF_PAGE_LIMIT: break
            text += page.get_text("text") + "\n"
        doc.close()
    except Exception as e:
        print(f"  - Error reading PDF with fitz {file_path}: {e}")
        return ""

    if len(text.strip()) < 100:
        print(f"  - Short text detected. Attempting OCR fallback for {os.path.basename(file_path)}...")
        try:
            pages = convert_from_path(file_path, last_page=PDF_PAGE_LIMIT)
            ocr_text = ""
            for i, page_image in enumerate(pages):
                ocr_text += pytesseract.image_to_string(page_image, lang="fra+ara") + "\n"
            return ocr_text.strip()
        except Exception as e:
            print(f"  - OCR fallback failed for {file_path}: {e}")
            return ""
    return text.strip()

def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        print(f"  - Error reading DOCX {file_path}: {e}")
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

# -----------------------------------------------------------------------------
# MAIN SCRIPT LOGIC
# -----------------------------------------------------------------------------
try:
    # --- PART 1: WEB SCRAPING AND DOWNLOADING ---
    print("\n--- Starting Part 1: Web Scraping ---")
    driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseHome")

    print("🔑 Logging in...")
    wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_login"))).send_keys("TARGETUPCONSULTING")
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password").send_keys(os.getenv("LOGIN_PASSWORD"))
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton").click()
    print("✅ Login successful.")

    print("🔍 Navigating to advanced search and setting filters...")
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
    time.sleep(2) # Allow time for table to reload

    rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')
    data = []
    for row in rows:
        try:
            ref = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
            objet = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", "")
            buyer = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", "")
            deadline = row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text.replace("\n", " ")
            link = row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")
            data.append({"reference": ref, "objet": objet, "acheteur": buyer, "date_limite": deadline, "download_page_url": link})
        except Exception:
            continue
    
    df = pd.DataFrame(data)
    excluded_words = ["construction", "installation", "recrutement", "travaux", "fourniture", "achat", "equipement", "maintenance", "works", "goods", "supply", "acquisition", "Recruitment", "nettoyage", "gardiennage"]
    df = df[~df['objet'].str.lower().str.contains('|'.join(excluded_words), na=False)]
    print(f"✅ Found {len(df)} relevant tenders after filtering.")

    # --- DOWNLOAD LOOP ---
    print("\n📥 Starting download loop...")
    links_to_process = df['download_page_url'].tolist()
    for i, link in enumerate(links_to_process):
        print(f"\n--- Processing link {i+1}/{len(links_to_process)} ---")
        try:
            driver.get(link)
            download_link = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce")))
            driver.execute_script("arguments[0].click();") # Use JS click for reliability
            
            checkbox = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions")))
            driver.execute_script("arguments[0].click();")
            
            driver.find_element(By.ID, "ctl0_CONTENU_PAGE_validateButton").click()
            
            final_download_button = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload")))
            final_download_button.click()
            print(f"✅ Download initiated. Waiting for file to save...")
            time.sleep(10) # Wait for download to finish
        except Exception as e:
            error_filename = f"error_page_{i+1}.png"
            print(f"⚠️ Could not download for link {link}. Saving screenshot to {error_filename}")
            driver.save_screenshot(error_filename)
            with open(f"error_page_{i+1}.html", "w", encoding="utf-8") as f: f.write(driver.page_source)
            traceback.print_exc()
            continue
    
    print("\n🎯 Download loop finished.")

    # --- PART 2: PROCESS DOWNLOADED FILES ---
    print("\n--- Starting Part 2: File Processing ---")

    print("Step 2.1: Unzipping all downloaded .zip files...")
    for root, _, files in os.walk(download_dir):
        for f in files:
            if f.lower().endswith(".zip"):
                extract_from_zip(os.path.join(root, f))
    
    print("\nStep 2.2: Extracting text from all files...")
    tender_results = []
    # Assumes download names are somewhat related to reference numbers, or we process all available folders
    for item_name in os.listdir(download_dir):
        item_path = os.path.join(download_dir, item_name)
        if not os.path.isdir(item_path): continue

        print(f"\nProcessing folder: {item_name}")
        merged_text = ""
        for root, _, files in os.walk(item_path):
            for f in files:
                if 'cps' in f.lower():
                    print(f"  -> Skipping file containing 'cps': {f}")
                    continue

                file_path = os.path.join(root, f)
                ext = os.path.splitext(f)[1].lower()
                text = ""

                if ext == ".pdf": text = extract_text_from_pdf(file_path)
                elif ext == ".docx": text = extract_text_from_docx(file_path)

                if text.strip():
                    merged_text += f"\n\n--- Content from: {f} ---\n{text}"
        
        if merged_text.strip():
            tender_results.append({"tender_folder": item_name, "merged_text": merged_text.strip()})
    
    df1 = pd.DataFrame(tender_results)

    # --- PART 3: MERGE AND SAVE RESULTS ---
    print("\n--- Starting Part 3: Merging data and saving to CSV ---")
    if not df.empty and not df1.empty:
        # This basic merge assumes the first scraped result corresponds to the first processed folder.
        # A more robust solution would match based on reference number in the folder/file names.
        df['id'] = range(len(df))
        df1['id'] = range(len(df1))
        
        merged_df = pd.merge(df, df1, on="id", how="inner")
        merged_df = merged_df.drop(columns=['id', 'reference']) # Drop temporary ID and redundant ref

        output_csv_path = "tender_results.csv"
        merged_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
        print(f"✅ Data successfully merged and saved to {output_csv_path}")
    else:
        print("⚠️ No data to process or merge. Saving initial scrape results only.")
        df.to_csv("tender_results.csv", index=False, encoding='utf-8-sig')


finally:
    print("\nQuitting WebDriver...")
    driver.quit()
    print("🎉 Script finished successfully.")
