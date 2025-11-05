import os
import time
import random
import zipfile
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException

# -----------------------------
# CONFIGURATION
# -----------------------------
DOWNLOAD_DIR = "/home/runner/work/downloads"
ZIP_DIR = os.path.join(DOWNLOAD_DIR, "zip_files")
UNZIP_DIR = os.path.join(DOWNLOAD_DIR, "unzipped")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(ZIP_DIR, exist_ok=True)
os.makedirs(UNZIP_DIR, exist_ok=True)

BASE_URL = "https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseHome"
USERNAME = "TARGETUPCONSULTING"  # replace with your test login
PASSWORD = "pgwr00jPD@"          # replace with your test password

# -----------------------------
# SELENIUM SETUP
# -----------------------------
options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
prefs = {
    "download.default_directory": ZIP_DIR,
    "download.prompt_for_download": False,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service("/usr/local/bin/chromedriver"), options=options)
wait = WebDriverWait(driver, 60)

# -----------------------------
# Helper: human typing
# -----------------------------
def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.05, 0.2))

# -----------------------------
# LOGIN
# -----------------------------
driver.get(BASE_URL)
time.sleep(2)

login_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_login")
password_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password")
ok_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton")

human_type(login_input, USERNAME)
time.sleep(random.uniform(0.5, 1))
human_type(password_input, PASSWORD)
time.sleep(random.uniform(0.5, 1))
ok_button.click()
time.sleep(3)

# -----------------------------
# Navigate to advanced search
# -----------------------------
driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons")
time.sleep(2)

# Set date filter to yesterday
date_input = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_dateMiseEnLigneCalculeStart")
yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
date_input.clear()
for char in yesterday:
    date_input.send_keys(char)
    time.sleep(random.uniform(0.05, 0.2))

time.sleep(1)
search_button = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche")
driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
search_button.click()
time.sleep(3)

# -----------------------------
# Set page size
# -----------------------------
dropdown = Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop"))
dropdown.select_by_value("500")
time.sleep(2)

# -----------------------------
# Scrape table rows
# -----------------------------
rows = driver.find_elements(By.XPATH, '//table[@class="table-results"]/tbody/tr')
data = []

for row in rows:
    try:
        ref = row.find_element(By.CSS_SELECTOR, '.col-450 .ref').text
        objet = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", "")
        buyer = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", "")
        lieux = row.find_element(By.XPATH, './/div[contains(@id,"panelBlocLieuxExec")]').text.replace("\n", ", ")
        deadline = row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text
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
        print(f"Error extracting row: {e}")

df = pd.DataFrame(data)

# -----------------------------
# DOWNLOAD ZIPs
# -----------------------------
for link in df['first_button_url']:
    try:
        driver.get(link)
        time.sleep(3)

        download_link = wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce")))
        driver.execute_script("arguments[0].scrollIntoView(true);", download_link)
        download_link.click()
        time.sleep(5)  # wait for download to complete
    except Exception as e:
        print(f"Download error for {link}: {e}")

driver.quit()

# -----------------------------
# UNZIP and prepare CSV
# -----------------------------
for f in os.listdir(ZIP_DIR):
    if f.lower().endswith(".zip"):
        zip_path = os.path.join(ZIP_DIR, f)
        extract_to = os.path.join(UNZIP_DIR, os.path.splitext(f)[0])
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)

# -----------------------------
# Save CSV
# -----------------------------
csv_path = os.path.join(DOWNLOAD_DIR, "tenders.csv")
df.to_csv(csv_path, index=False)
print(f"✅ CSV saved at {csv_path}")
