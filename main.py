import os
import time
import re
import shutil
import zipfile
import subprocess
import traceback
import unicodedata
import pandas as pd
from datetime import datetime, timedelta

# PDF / OCR / DOC
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import docx

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException

# Requests for sending to n8n
import requests

# -------------------------------------------------------------------------
# INITIAL SETUP
# -------------------------------------------------------------------------
print("🚀 Initializing configuration for GitHub Actions environment...")
download_dir = os.path.join(os.getcwd(), "downloads_temp")
os.makedirs(download_dir, exist_ok=True)

options = webdriver.ChromeOptions()
user_agent = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
)
options.add_argument(f"user-agent={user_agent}")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
}
options.add_experimental_option("prefs", prefs)

service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)
driver.set_page_load_timeout(60)
print("✅ WebDriver initialized successfully.")

# -------------------------------------------------------------------------
# HELPER FUNCTIONS
# -------------------------------------------------------------------------
PDF_PAGE_LIMIT = 10

def clean_extracted_text(text):
    """Cleans and prettifies extracted text."""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\n{2,}", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"Page\s*\d+\s*/\s*\d+", "", text, flags=re.IGNORECASE)
    text = re.sub(r"[\u0000-\u001f]+", "", text)
    text = re.sub(r"([.?!])\s+([a-z])", lambda m: m.group(1) + " " + m.group(2).upper(), text)

    cleaned_lines = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if len(line) < 3 and not re.search(r"\d", line):
            continue
        cleaned_lines.append(line)

    pretty_text = "\n".join(cleaned_lines)
    pretty_text = re.sub(r"\n{3,}", "\n\n", pretty_text)
    return pretty_text.strip()


def extract_text_from_pdf(file_path):
    """Extracts readable text from PDF with OCR fallback."""
    text = ""
    try:
        doc = fitz.open(file_path)
        page_count = min(len(doc), PDF_PAGE_LIMIT)
        for i in range(page_count):
            page_text = doc[i].get_text("text")
            if page_text:
                text += page_text + "\n"
        doc.close()
    except Exception:
        text = ""
    
    if len(text.strip()) < 50:
        print("    - Short text from PDF, attempting OCR fallback...")
        try:
            pages = convert_from_path(file_path, last_page=PDF_PAGE_LIMIT)
            ocr_text = ""
            for page_image in pages:
                ocr_text += pytesseract.image_to_string(page_image, lang="fra+ara+eng") + "\n"
            text = ocr_text
        except Exception as e:
            print(f"    - ⚠️ OCR failed. Error: {e}")
            text = ""
    return clean_extracted_text(text)


def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        text = "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
        return clean_extracted_text(text)
    except Exception:
        return ""


def extract_text_from_doc(file_path):
    try:
        process = subprocess.Popen(["antiword", file_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, _ = process.communicate()
        text = stdout.decode("utf-8", errors="ignore")
        return clean_extracted_text(text)
    except Exception as e:
        print(f"    - ⚠️ Antiword failed for .doc file. Error: {e}")
        return ""


def extract_from_zip(file_path):
    try:
        extract_to = os.path.splitext(file_path)[0]
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(extract_to)
        return extract_to
    except Exception as e:
        print(f"  - Failed to unzip {file_path}: {e}")
        return None


def clear_download_directory():
    for item in os.listdir(download_dir):
        path = os.path.join(download_dir, item)
        try:
            if os.path.isfile(path) or os.path.islink(path):
                os.unlink(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)
        except Exception as e:
            print(f"⚠️ Failed to delete {path}: {e}")


def wait_for_download_complete(timeout=90):
    """Waits for Chrome to finish downloading a file."""
    seconds = 0
    while seconds < timeout:
        downloading = any(
            f.endswith(".crdownload") or f.startswith(".com.google.Chrome")
            for f in os.listdir(download_dir)
        )
        if not downloading:
            files = [
                f for f in os.listdir(download_dir)
                if not (f.endswith(".crdownload") or f.startswith(".com.google.Chrome"))
            ]
            if files:
                return os.path.join(download_dir, files[0])
        time.sleep(1)
        seconds += 1
    return None

# -------------------------------------------------------------------------
# MAIN SCRIPT LOGIC
# -------------------------------------------------------------------------
df = pd.DataFrame()

try:
    print("\n--- Starting Part 1: Scraping tender metadata ---")
    URL1 = os.getenv("URL1")
    driver.get(URL1)
    print("🔑 Logging in...")
    wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_login"))).send_keys(os.getenv("USERNAME"))
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_password").send_keys(os.getenv("PASSWORD"))
    driver.find_element(By.ID, "ctl0_CONTENU_PAGE_authentificationButton").click()
    print("✅ Login successful.")

    print("🔍 Navigating to search and applying filters...")
    time.sleep(3)
    URL2 = os.getenv("URL2")
    driver.get(URL2)
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
            data.append({
                "reference": row.find_element(By.CSS_SELECTOR, ".col-450 .ref").text,
                "objet": row.find_element(By.XPATH, './/div[contains(@id,"panelBlocObjet")]').text.replace("Objet : ", ""),
                "acheteur": row.find_element(By.XPATH, './/div[contains(@id,"panelBlocDenomination")]').text.replace("Acheteur public : ", ""),
                "lieux_execution": row.find_element(By.XPATH, './/div[contains(@id,"panelBlocLieuxExec")]').text.replace("\n", ", "),
                "date_limite": row.find_element(By.XPATH, './/td[@headers="cons_dateEnd"]').text.replace("\n", " "),
                "download_page_url": row.find_element(By.XPATH, './/td[@class="actions"]//a[1]').get_attribute("href")
            })
        except Exception:
            continue

    df = pd.DataFrame(data)
    excluded_words = [
        # French / English equivalents
        "construction", "construction",
        "installation", "installation",
        "recrutement", "recruitment",
        "travaux", "works",
        "fourniture", "supply",
        "achat", "purchase",
        "equipement", "equipment",
        "maintenance", "maintenance",
        "works", "works",
        "goods", "goods",
        "supply", "supply",
        "acquisition", "acquisition",
        "Recruitment", "recruitment",
        "nettoyage", "cleaning",
        "gardiennage", "guarding",
        "archives", "archives", "archivage",
        "Equipment", "equipment",
        "ÉQUIPEMENT", "equipment",
        "équipement", "equipment",
        "construire", "build",
        "recrute", "recruits"
    ]


    if not df.empty:
        df_filtered = df[~df["objet"].str.lower().str.contains("|".join(excluded_words), na=False)].reset_index(drop=True)
    else:
        df_filtered = df

    print(f"✅ Found {len(df_filtered)} relevant tenders after filtering.")

    print("\n--- Starting Part 2: Processing tenders ---\n")
    all_processed_tenders = []

    for index, row in df_filtered.iterrows():
        print(f"--- Processing Tender {index + 1}/{len(df_filtered)} | Ref: {row['reference']} ---")
        try:
            clear_download_directory()
            driver.get(row["download_page_url"])

            dce_link = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dce_link)
            time.sleep(1)
            dce_link.click()

            checkbox = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions")))
            checkbox.click()

            validate_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_validateButton")))
            validate_btn.click()

            final_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload")))
            final_btn.click()

            print("  - Download command sent. Waiting for file...")
            downloaded = wait_for_download_complete()

            merged_text = ""
            if not downloaded:
                print("  - ⚠️ Download failed or timed out.")
                merged_text = "Error: download failed."
            else:
                print(f"  - ✅ Download complete: {os.path.basename(downloaded)}")
                file_paths = []
                if downloaded.lower().endswith(".zip"):
                    unzip_dir = extract_from_zip(downloaded)
                    if unzip_dir:
                        for r, _, files in os.walk(unzip_dir):
                            for f in files:
                                file_paths.append(os.path.join(r, f))
                else:
                    file_paths.append(downloaded)

                print(f"  - Found {len(file_paths)} file(s) to process.")
                for fpath in file_paths:
                    fname = os.path.basename(fpath)
                    print(f"    -> Processing file: {fname}")

                    if "cps" in fname.lower():
                        print("       - Skipping CPS file.")
                        continue

                    ext = os.path.splitext(fname)[1].lower()
                    text = ""
                    if ext == ".pdf":
                        text = extract_text_from_pdf(fpath)
                    elif ext == ".docx":
                        text = extract_text_from_docx(fpath)
                    elif ext == ".doc":
                        text = extract_text_from_doc(fpath)

                    if text:
                        merged_text += f"\n\n====================\n--- Content from file: {fname} ---\n{text}\n====================\n"
                    else:
                        print("       - No text extracted.")

            tender_payload = row.to_dict()
            tender_payload["merged_text"] = merged_text.strip() or "No relevant text could be extracted."
            # tender_payload.pop("download_page_url", None)

            webhook = os.getenv("N8N_WEBHOOK_URL")
            if webhook:
                print("  - Sending to n8n...")
                try:
                    resp = requests.post(webhook, json=tender_payload, timeout=30)
                    if resp.status_code == 200:
                        print("  - ✅ Successfully sent!")
                        time.sleep(10)
                    else:
                        print(f"  - ❌ Failed to send (status {resp.status_code})")
                except Exception as e:
                    print(f"  - ❌ Error sending to n8n: {e}")

            all_processed_tenders.append(tender_payload)

        except Exception as e:
            print(f"  - ⚠️ Error on this tender: {type(e).__name__}")
            traceback.print_exc()
            continue

finally:
    print("\n--- Finalizing Script ---")
    if "all_processed_tenders" in locals() and all_processed_tenders:
        df_out = pd.DataFrame(all_processed_tenders)
        out_path = os.path.join(os.getcwd(), "tender_results_summary.csv")
        df_out.to_csv(out_path, index=False, encoding="utf-8-sig")
        print(f"✅ Saved summary for {len(df_out)} tenders to {out_path}")
    else:
        print("ℹ️ No tenders processed.")

    print("Quitting WebDriver...")
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
