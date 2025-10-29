import os
import time
import zipfile
import random
import shutil
import tempfile
from pathlib import Path
import pandas as pd
import PyPDF2
import docx
import openpyxl
import chardet
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC

# ------------------------
# Configuration
options = webdriver.ChromeOptions()
prefs = {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
# options.add_argument("--headless=new")  # Uncomment to run headless

service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 10)

# ------------------------
# TEMPORARY FOLDER
temp_dir = tempfile.mkdtemp()

# ------------------------
# Helper Functions
def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.08, 0.25))

def extract_text_from_file(file_path):
    text = ""
    file_path = Path(file_path)
    if file_path.suffix.lower() == ".pdf":
        try:
            with open(file_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                text = "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception as e:
            print(f"PDF read error: {e}")
    elif file_path.suffix.lower() == ".docx":
        try:
            doc = docx.Document(file_path)
            text = "\n".join([p.text for p in doc.paragraphs])
        except Exception as e:
            print(f"DOCX read error: {e}")
    elif file_path.suffix.lower() == ".xlsx":
        try:
            xls = pd.ExcelFile(file_path)
            sheets_text = []
            for sheet in xls.sheet_names:
                df_sheet = xls.parse(sheet).astype(str).apply(lambda x: " ".join(x), axis=1)
                sheets_text.append(df_sheet.str.cat(sep="\n"))
            text = "\n".join(sheets_text)
        except Exception as e:
            print(f"XLSX read error: {e}")
    elif file_path.suffix.lower() == ".csv":
        try:
            with open(file_path, "rb") as f:
                result = chardet.detect(f.read())
            df_csv = pd.read_csv(file_path, encoding=result['encoding'])
            text = df_csv.astype(str).apply(lambda x: " ".join(x), axis=1).str.cat(sep="\n")
        except Exception as e:
            print(f"CSV read error: {e}")
    return text

def process_download(file_path):
    file_path = Path(file_path)
    extracted_texts = []
    if zipfile.is_zipfile(file_path):
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(temp_dir)
            for f in zip_ref.namelist():
                extracted_texts.append(extract_text_from_file(Path(temp_dir)/f))
    else:
        extracted_texts.append(extract_text_from_file(file_path))
    return "\n".join(extracted_texts)

# ------------------------
# LOGIN AND NAVIGATION
driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseHome")
time.sleep(2)

login_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_login")
password_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password")
ok_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton")

email = "TARGETUPCONSULTING"
password = "pgwr00jPD@"

human_type(login_input, email)
time.sleep(random.uniform(0.5, 1.2))
human_type(password_input, password)
time.sleep(random.uniform(0.5, 1.5))
ok_button.click()

# SEARCH
driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons")
time.sleep(2)

date_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")
yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
date_input.clear()
for char in yesterday:
    date_input.send_keys(char)
    time.sleep(random.uniform(0.08, 0.2))
time.sleep(random.uniform(0.5, 1))

search_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche")
driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", search_button)
time.sleep(random.uniform(0.5, 1.2))
search_button.click()
time.sleep(2)

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
            "reference": ref,
            "objet": objet,
            "acheteur": buyer,
            "lieux_execution": lieux,
            "date_limite": deadline,
            "first_button_url": first_button
        })
    except Exception as e:
        print(f"Row error: {e}")

df = pd.DataFrame(data)

# ------------------------
# DOWNLOAD & EXTRACT TEXT
for idx, link in enumerate(df['first_button_url']):
    driver.get(link)
    time.sleep(2)
    try:
        download_link = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce")
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", download_link)
        time.sleep(random.uniform(0.5, 1.2))
        download_link.click()
        time.sleep(2)
        
        checkbox = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions")
        if not checkbox.is_selected():
            checkbox.click()
        valider_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_validateButton")
        valider_button.click()
        time.sleep(2)
        
        download_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload")
        download_button.click()
        time.sleep(3)
        
        # Move downloaded file to temp folder
        default_download_dir = Path.home() / "Downloads"
        downloaded_file = max(default_download_dir.glob("*"), key=lambda f: f.stat().st_mtime)
        shutil.move(str(downloaded_file), Path(temp_dir)/downloaded_file.name)
        
        # Process file
        df.loc[idx, "extracted_text"] = process_download(Path(temp_dir)/downloaded_file.name)
        
    except Exception as e:
        print(f"Download error: {e}")

# ------------------------
# SAVE OUTPUT
df.to_csv("output.csv", index=False)
print("Done! Extracted text saved to output.csv")

# CLEANUP
driver.quit()
shutil.rmtree(temp_dir)
