import os
import time
import zipfile
import random
from datetime import datetime, timedelta
import pandas as pd
import mimetypes
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    ElementClickInterceptedException
)
from PyPDF2 import PdfReader
from docx import Document
import openpyxl
import chardet

# ---------------- Configuration ----------------
TMP_DOWNLOAD_DIR = os.environ.get("TMP_DOWNLOAD_DIR", "downloads")
os.makedirs(TMP_DOWNLOAD_DIR, exist_ok=True)

options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": os.path.abspath(TMP_DOWNLOAD_DIR),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
# options.add_argument("--headless=new")  # Uncomment for headless

service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 10)

# ---------------- Helper Functions ----------------
def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.08, 0.25))

def extract_text_from_file(path):
    text = ""
    ext = path.split('.')[-1].lower()
    try:
        if ext == "pdf":
            reader = PdfReader(path)
            for page in reader.pages:
                text += page.extract_text() or ""
        elif ext == "docx":
            doc = Document(path)
            for p in doc.paragraphs:
                text += p.text + "\n"
        elif ext in ["xls", "xlsx"]:
            wb = openpyxl.load_workbook(path, data_only=True)
            for sheet in wb:
                for row in sheet.iter_rows(values_only=True):
                    text += " ".join([str(c) if c else "" for c in row]) + "\n"
        elif ext in ["txt", "csv"]:
            with open(path, "rb") as f:
                raw = f.read()
                enc = chardet.detect(raw)['encoding'] or 'utf-8'
            with open(path, "r", encoding=enc, errors="ignore") as f:
                text = f.read()
        elif ext == "zip":
            with zipfile.ZipFile(path, 'r') as zip_ref:
                for file_name in zip_ref.namelist():
                    zip_ref.extract(file_name, TMP_DOWNLOAD_DIR)
                    full_path = os.path.join(TMP_DOWNLOAD_DIR, file_name)
                    text += extract_text_from_file(full_path)
    except Exception as e:
        print(f"Error extracting {path}: {e}")
    return text

def click_element(el):
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", el)
    time.sleep(random.uniform(0.3, 0.8))
    el.click()

# ---------------- Login ----------------
driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseHome")
time.sleep(2)
login_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_login")
password_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password")
ok_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton")

email = "TARGETUPCONSULTING"
password = "pgwr00jPD@"

human_type(login_input, email)
time.sleep(random.uniform(0.5,1.2))
human_type(password_input, password)
time.sleep(random.uniform(0.5,1.5))
click_element(ok_button)

# ---------------- Search ----------------
driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons")
time.sleep(2)
date_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")
yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
date_input.clear()
for char in yesterday:
    date_input.send_keys(char)
    time.sleep(random.uniform(0.08, 0.2))

search_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche")
click_element(search_button)
time.sleep(2)

Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")).select_by_value("500")
time.sleep(2)

# ---------------- Scrape Rows ----------------
data = []
num_rows = len(driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr'))
last_screenshot_time = time.time()

for i in range(num_rows):
    try:
        rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')
        row = rows[i]

        ref = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
        objet = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", "")
        buyer = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", "")
        lieux = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocLieuxExec")]').text.replace("\n", ", ")
        deadline = row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text.replace("\n", " ")
        first_button = row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")

        data.append({
            "reference": ref,
            "objet": objet,
            "acheteur": buyer,
            "lieux_execution": lieux,
            "date_limite": deadline,
            "first_button_url": first_button,
            "extracted_text": ""
        })

        # Screenshot every second
        if time.time() - last_screenshot_time >= 1:
            driver.save_screenshot(os.path.join(TMP_DOWNLOAD_DIR, f"progress_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"))
            last_screenshot_time = time.time()

    except Exception as e:
        print(f"Error: Row {i+1} failed: {e}")
        driver.save_screenshot(os.path.join(TMP_DOWNLOAD_DIR, f"row_{i+1}_error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"))

# ---------------- Download & Extract ----------------
for idx, row_data in enumerate(data):
    try:
        driver.get(row_data["first_button_url"])
        time.sleep(2)

        download_link = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce")
        click_element(download_link)
        time.sleep(2)

        checkbox = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions")
        if not checkbox.is_selected():
            click_element(checkbox)
        valider_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_validateButton")
        click_element(valider_button)
        time.sleep(2)

        download_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload")
        click_element(download_button)
        time.sleep(3)

        # Pick the latest downloaded file
        files = os.listdir(TMP_DOWNLOAD_DIR)
        files = [os.path.join(TMP_DOWNLOAD_DIR, f) for f in files]
        latest_file = max(files, key=os.path.getctime)
        extracted_text = extract_text_from_file(latest_file)
        data[idx]["extracted_text"] = extracted_text

        print(f"Processed row {idx+1}, file: {latest_file}")

    except Exception as e:
        print(f"Download/extract failed for row {idx+1}: {e}")
        driver.save_screenshot(os.path.join(TMP_DOWNLOAD_DIR, f"download_error_{idx+1}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"))

# ---------------- Save CSV ----------------
df = pd.DataFrame(data)
csv_file = os.path.join(TMP_DOWNLOAD_DIR, f"marchespublics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
df.to_csv(csv_file, index=False)
print(f"CSV saved: {csv_file}")

driver.quit()
