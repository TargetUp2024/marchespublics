import os
import time
import zipfile
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
import random

# Additional imports for file parsing
import PyPDF2
import docx

# ------------------------
# CONFIGURATION
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--headless=new")  # headless for GitHub Actions
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 10)

print("[INFO] Browser initialized in headless mode...")

# ------------------------
# LOGIN
driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseHome")
time.sleep(2)

def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.08, 0.25))

login_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_login")
password_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password")
ok_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton")

email = "TARGETUPCONSULTING"
password = "pgwr00jPD@"

print("[INFO] Typing credentials...")
human_type(login_input, email)
time.sleep(random.uniform(0.5,1))
human_type(password_input, password)
time.sleep(random.uniform(0.5,1))
ok_button.click()
time.sleep(2)
print("[INFO] Logged in successfully.")

# ------------------------
# SEARCH FOR YESTERDAY
driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons")
time.sleep(2)

date_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")
yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
date_input.clear()
for char in yesterday:
    date_input.send_keys(char)
    time.sleep(random.uniform(0.08,0.2))

search_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche")
search_button.click()
time.sleep(2)
print(f"[INFO] Searching for tenders posted on {yesterday}...")

dropdown = Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop"))
dropdown.select_by_value("10")
time.sleep(2)

rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')
print(f"[INFO] Found {len(rows)} rows on the page.")

data = []

# ------------------------
# Helper function to extract text from various file types
def extract_text_from_file(file_path):
    text = ""
    if zipfile.is_zipfile(file_path):
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(DOWNLOAD_DIR)
            for f in zip_ref.namelist():
                path = os.path.join(DOWNLOAD_DIR, f)
                text += extract_text_from_file(path) + "\n"
    elif file_path.endswith(".txt"):
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()
    elif file_path.endswith(".pdf"):
        try:
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    text += page.extract_text() + "\n"
        except:
            text = ""
    elif file_path.endswith(".docx"):
        try:
            doc = docx.Document(file_path)
            text = "\n".join([p.text for p in doc.paragraphs])
        except:
            text = ""
    elif file_path.endswith(".xlsx") or file_path.endswith(".csv"):
        try:
            df_file = pd.read_excel(file_path) if file_path.endswith(".xlsx") else pd.read_csv(file_path)
            text = df_file.astype(str).apply(lambda x: ' '.join(x), axis=1).str.cat(sep='\n')
        except:
            text = ""
    return text

# ------------------------
# PROCESS ROWS & DOWNLOAD FILES
for idx, row in enumerate(rows, start=1):
    try:
        ref = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
        objet = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", "")
        buyer = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", "")
        lieux = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocLieuxExec")]').text.replace("\n", ", ")
        deadline = row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text.replace("\n"," ")
        first_button = row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")

        print(f"[INFO] Processing row {idx}/{len(rows)}: {ref}")

        # Go to download page
        driver.get(first_button)
        time.sleep(2)

        checkbox = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions")
        if not checkbox.is_selected():
            checkbox.click()
            time.sleep(0.5)
        driver.find_element(By.ID, "ctl0_CONTENU_PAGE_validateButton").click()
        time.sleep(1)
        driver.find_element(By.ID, "ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload").click()

        time.sleep(5 + random.randint(1,3))  # wait for download

        files = os.listdir(DOWNLOAD_DIR)
        downloaded_file = max([os.path.join(DOWNLOAD_DIR,f) for f in files], key=os.path.getctime)
        print(f"[INFO] Downloaded file: {downloaded_file}")

        extracted_texts = extract_text_from_file(downloaded_file)

        data.append({
            "reference": ref,
            "objet": objet,
            "acheteur": buyer,
            "lieux_execution": lieux,
            "date_limite": deadline,
            "first_button_url": first_button,
            "dce_text": extracted_texts
        })

    except Exception as e:
        print(f"[ERROR] Row {idx} failed: {e}")

# ------------------------
# SAVE TO CSV
df = pd.DataFrame(data)
csv_path = os.path.join(os.getcwd(), f"marchespublics_{yesterday.replace('/','-')}.csv")
df.to_csv(csv_path, index=False, encoding='utf-8-sig')
print(f"[INFO] CSV saved at {csv_path}")
print("[INFO] Done.")
