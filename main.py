import os, time, csv, zipfile, pandas as pd
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pdfminer.high_level import extract_text
import docx

BASE_URL = "https://www.marchespublics.gov.ma/"

DATA_DIR = Path("scraped_data")
ZIP_DIR = DATA_DIR / "zip_files"
EXTRACT_DIR = DATA_DIR / "unzipped"
CSV_PATH = DATA_DIR / "tenders.csv"

for folder in [DATA_DIR, ZIP_DIR, EXTRACT_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

USERNAME = "TARGETUPCONSULTING"
PASSWORD = "pgwr00jPD@"

options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

def extract_file_text(file_path):
    f = str(file_path)
    if f.endswith(".pdf"):
        try: return extract_text(f)
        except: return "PDF error"
    if f.endswith(".docx"):
        try:
            d = docx.Document(f)
            return "\n".join([p.text for p in d.paragraphs])
        except: return "DOCX error"
    if f.endswith(".txt"):
        return open(f, encoding="utf-8", errors="ignore").read()
    return ""

print("Opening site...")
driver.get(BASE_URL)
wait.until(EC.element_to_be_clickable((By.ID, "login-button"))).click()
wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(USERNAME)
wait.until(EC.presence_of_element_located((By.ID, "password"))).send_keys(PASSWORD + Keys.RETURN)

time.sleep(3)
print("Going to search page...")
driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons")

time.sleep(3)
rows = driver.find_elements(By.CSS_SELECTOR, ".table > tbody > tr")

records = []

for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")
    tender_ref = cols[0].text.strip()
    tender_title = cols[1].text.strip()
    
    # click tender
    link = cols[1].find_element(By.TAG_NAME, "a")
    tender_url = link.get_attribute("href")
    driver.execute_script("window.open(arguments[0]);", tender_url)
    driver.switch_to.window(driver.window_handles[-1])

    time.sleep(3)

    text_accumulated = ""

    try:
        download_btn = driver.find_element(By.XPATH, "//a[contains(@href, 'telechargerDce')]")
        zip_url = download_btn.get_attribute("href")

        driver.get(zip_url)
        time.sleep(5)

        # get latest downloaded ZIP from temp
        zip_files = list(Path("/home/runner/").rglob("*.zip"))  # GitHub actions default download dir
        if zip_files:
            latest_zip = max(zip_files, key=os.path.getctime)
            local_zip = ZIP_DIR / f"{tender_ref}.zip"
            os.rename(latest_zip, local_zip)

            with zipfile.ZipFile(local_zip, 'r') as z:
                z.extractall(EXTRACT_DIR / tender_ref)

            for f in (EXTRACT_DIR / tender_ref).rglob("*"):
                if f.is_file():
                    text_accumulated += "\n\n" + extract_file_text(f)

    except Exception as e:
        text_accumulated = "No documents or error"

    records.append({
        "Reference": tender_ref,
        "Title": tender_title,
        "Extracted Text": text_accumulated.strip()
    })

    driver.close()
    driver.switch_to.window(driver.window_handles[0])

driver.quit()

df = pd.DataFrame(records)
df.to_csv(CSV_PATH, index=False, encoding="utf-8")

print("✅ DONE — CSV saved at:")
print(CSV_PATH)
