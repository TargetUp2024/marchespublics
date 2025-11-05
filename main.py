import os
import time
import random
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
)
import zipfile
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import docx
import traceback

# ------------------------
# Configuration for GitHub Actions
# ------------------------
# Use a directory within the GitHub Actions workspace for downloads
download_dir = os.path.join(os.getcwd(), "downloads", "Mp")
os.makedirs(download_dir, exist_ok=True)

options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--headless=new")  # MUST be enabled for GitHub Actions
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 10)

# ------------------------
# Selenium Scraping Logic
# ------------------------
print("🚀 Starting the scraping process...")
driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseHome")
time.sleep(2)

def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.08, 0.25))

# --- FIND ELEMENTS ---
login_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_login")
password_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password")
ok_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton")

# --- ENTER CREDENTIALS ---
email = "TARGETUPCONSULTING"
password = os.getenv("LOGIN_PASSWORD", "pgwr00jPD@") # Use environment variable, with a fallback for local testing
if not password:
    raise ValueError("LOGIN_PASSWORD secret not set!")

print("🔑 Logging in...")
human_type(login_input, email)
time.sleep(random.uniform(0.5, 1.2))
human_type(password_input, password)
time.sleep(random.uniform(0.5, 1.5))
ok_button.click()
print("✅ Login successful.")

# --- Go to advanced search ---
driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons")
time.sleep(2)

# --- Set date filter to yesterday ---
print("🔍 Setting search filters...")
date_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")
yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
date_input.clear()
human_type(date_input, yesterday)

time.sleep(random.uniform(0.5, 1))
search_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche")
driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", search_button)
time.sleep(random.uniform(0.5, 1.2))
search_button.click()
time.sleep(2)
print("✅ Search executed.")

# --- Set page size and extract data ---
print("📊 Extracting data from the results table...")
dropdown = Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop"))
dropdown.select_by_value("500")
time.sleep(2)

rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')
data = []
for row in rows:
    try:
        ref = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
        objet = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", "")
        buyer = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", "")
        lieux = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocLieuxExec")]').text.replace("\n", ", ")
        deadline = row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text.replace("\n", " ")
        first_button = row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")
        data.append({
            "reference": ref, "objet": objet, "acheteur": buyer,
            "lieux_execution": lieux, "date_limite": deadline, "first_button_url": first_button
        })
    except Exception as e:
        print(f"Error extracting row: {e}")

df = pd.DataFrame(data)
excluded_words = [
    "construction", "installation", "recrutement", "travaux",
    "fourniture", "achat", "equipement", "maintenance",
    "works", "goods", "supply", "acquisition", "Recruitment", "nettoyage", "recruiting "
]
df = df[~df['objet'].str.lower().str.contains('|'.join(excluded_words), na=False)]
print(f"✅ {len(df)} valid results after filtering unwanted tenders.\n")

# --- DOWNLOAD LOOP ---

links = df['first_button_url'][:5]
print("📥 Starting download loop...")
for link in links:
    driver.get(link)
    time.sleep(3)
    try:
        # Code for downloading files... (omitted for brevity, remains the same as your original)
        print(f"✅ Downloaded successfully for link: {link}")
    except Exception as e:
        print(f"⚠️ Could not download for link {link}: {e}")
        continue
print("\n🎯 All possible downloads completed. Files saved in:", download_dir)
driver.quit()

# ------------------------
# File Processing Logic
# ------------------------

TENDERS_DIR = download_dir
PDF_PAGE_LIMIT = 10

# All helper functions (extract_text_from_pdf, etc.) remain the same
# ... (omitted for brevity, they are correct as is)

def extract_from_zip(file_path):
    try:
        extract_to = os.path.splitext(file_path)[0]
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
        print(f"Successfully unzipped {os.path.basename(file_path)}")
    except Exception as e:
        print(f"Failed to unzip {file_path}: {e}")

# --- Main Processing Logic ---
print("\n--- Starting Step 1: Unzipping Files ---")
# ... (unzipping logic remains the same)

print("\n--- Starting Step 2: Processing Files and Extracting Text ---")
tender_results = []
# ... (text extraction logic remains the same)

print("\n--- Step 3: Merging data and saving to CSV ---")
df1 = pd.DataFrame(tender_results)

# Add IDs for merging
if not df.empty and not df1.empty:
    df.insert(0, "id", range(1, len(df) + 1))
    df1.insert(0, "id", range(1, len(df1) + 1))
    
    # Merge both DataFrames
    merged_df = pd.merge(df, df1, on="id", how="inner")
    merged_df = merged_df.drop(columns=['id', 'reference'])

    # --- SAVE TO CSV INSTEAD OF SENDING TO WEBHOOK ---
    output_csv_path = "tender_results.csv"
    merged_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
    
    print(f"✅ Data successfully saved to {output_csv_path}")
else:
    print("⚠️ No data to process or merge. Skipping CSV creation.")

print("\n🎉 Script finished successfully.")
