import os
import time
import re
import shutil
import zipfile
import subprocess
import traceback
import unicodedata
import random
import requests
import pandas as pd
from datetime import datetime, timedelta

# PDF / OCR / DOC
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import docx

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# -----------------------------
# CONFIGURATION
# -----------------------------
print("🚀 Initializing configuration...")
download_dir = os.path.join(os.getcwd(), "downloads_temp")
os.makedirs(download_dir, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)

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
print("✅ WebDriver initialized.")

PDF_PAGE_LIMIT = 10

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def clean_extracted_text(text):
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
        try:
            pages = convert_from_path(file_path, last_page=PDF_PAGE_LIMIT)
            ocr_text = ""
            for page_image in pages:
                ocr_text += pytesseract.image_to_string(page_image, lang="fra+ara+eng") + "\n"
            text = ocr_text
        except Exception as e:
            print(f"⚠️ OCR failed for {file_path}: {e}")
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
        print(f"⚠️ Antiword failed for {file_path}: {e}")
        return ""


def extract_from_zip(file_path):
    try:
        extract_to = os.path.splitext(file_path)[0]
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(extract_to)
        return extract_to
    except Exception as e:
        print(f"⚠️ Failed to unzip {file_path}: {e}")
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
    seconds = 0
    while seconds < timeout:
        downloading = any(f.endswith(".crdownload") or f.startswith(".com.google.Chrome") for f in os.listdir(download_dir))
        if not downloading:
            files = [f for f in os.listdir(download_dir) if not f.endswith(".crdownload")]
            if files:
                return os.path.join(download_dir, files[0])
        time.sleep(1)
        seconds += 1
    return None

# -----------------------------
# MAIN SCRIPT
# -----------------------------
all_processed_tenders = []
df = pd.DataFrame()

try:
    print("\n--- Starting scraping ---")
    driver.get("https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons")
    time.sleep(2)

    # Step 1: Open "Définir" popup
    define_btn = wait.until(
        EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_AdvancedSearch_domaineActivite_linkDisplay"))
    )
    define_btn.click()

    wait.until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])

    # Step 2: Select Services checkbox and validate
    checkbox = wait.until(
        EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_repeaterCategorie_ctl2_idCategorie"))
    )
    checkbox.click()
    validate_btn = wait.until(
        EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_validateButton"))
    )
    validate_btn.click()
    driver.switch_to.window(driver.window_handles[0])
    print("✅ Services selected.")

    # Step 3: Set date filter to yesterday
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

    # Step 4: Set results per page
    wait.until(EC.presence_of_element_located((By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop")))
    dropdown = Select(driver.find_element(By.ID, "ctl0_CONTENU_PAGE_resultSearch_listePageSizeTop"))
    dropdown.select_by_value("500")
    time.sleep(2)

    # Step 5: Scrape table rows
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
            print(f"Error extracting row: {e}")

    df = pd.DataFrame(data)

    # Step 6: Filter unwanted tenders
    excluded_words = [
        "construction", "installation", "travaux",
        "fourniture", "achat", "equipement", "supply", "acquisition", "nettoyage"
    ]
    
    df = df[~df['objet'].str.lower().str.contains('|'.join(excluded_words), na=False)]
    print(f"✅ {len(df)} valid tenders after filtering.\n")

    # Step 7: Download loop
    fields = {
        "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_nom": "Lachhab",
        "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_prenom": "Anas",
        "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_email": "anas.lachhab@example.com"
    }

    for idx, row in df.iterrows():
        link = row['first_button_url']
        driver.get(link)
        time.sleep(3)
        try:
            download_link = wait.until(
                EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", download_link)
            download_link.click()

            # Fill form
            for fid, value in fields.items():
                input_field = wait.until(EC.presence_of_element_located((By.ID, fid)))
                input_field.clear()
                input_field.send_keys(value)

            # Accept terms
            checkbox = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions")
            if not checkbox.is_selected():
                checkbox.click()

            # Validate
            valider_button = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_validateButton")))
            valider_button.click()

            # Final download
            final_button = wait.until(
                EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload"))
            )
            final_button.click()
            print(f"✅ Download started for {link}")

            # Wait for file
            downloaded_file = wait_for_download_complete()
            if not downloaded_file:
                print("  - ⚠️ Download failed or timed out.")
                merged_text = "Error: download failed."
            else:
                file_paths = []
                if downloaded_file.lower().endswith(".zip"):
                    unzip_dir = extract_from_zip(downloaded_file)
                    if unzip_dir:
                        for r, _, files in os.walk(unzip_dir):
                            for f in files:
                                file_paths.append(os.path.join(r, f))
                else:
                    file_paths.append(downloaded_file)

                merged_text = ""
                for fpath in file_paths:
                    fname = os.path.basename(fpath)
                    ext = os.path.splitext(fname)[1].lower()
                    if "cps" in fname.lower():
                        continue
                    if ext == ".pdf":
                        merged_text += extract_text_from_pdf(fpath)
                    elif ext == ".docx":
                        merged_text += extract_text_from_docx(fpath)
                    elif ext == ".doc":
                        merged_text += extract_text_from_doc(fpath)

            tender_payload = row.to_dict()
            tender_payload["merged_text"] = merged_text.strip() or "No relevant text could be extracted."

            # Send to n8n
            webhook = os.getenv("N8N_WEBHOOK_URL")
            if webhook:
                try:
                    resp = requests.post(webhook, json=tender_payload, timeout=30)
                    if resp.status_code == 200:
                        print("  - ✅ Sent to n8n successfully")
                    else:
                        print(f"  - ❌ n8n error: {resp.status_code}")
                except Exception as e:
                    print(f"  - ❌ n8n exception: {e}")

            all_processed_tenders.append(tender_payload)

        except Exception as e:
            print(f"⚠️ Error processing tender {link}: {e}")
            traceback.print_exc()
            continue

finally:
    # Save summary
    if all_processed_tenders:
        df_out = pd.DataFrame(all_processed_tenders)
        out_path = os.path.join(os.getcwd(), "tender_results_summary.csv")
        df_out.to_csv(out_path, index=False, encoding="utf-8-sig")
        print(f"✅ Saved {len(df_out)} tenders to {out_path}")
    else:
        print("ℹ️ No tenders processed.")

    # Cleanup
    try:
        driver.quit()
    except Exception:
        pass
    if os.path.exists(download_dir):
        try:
            shutil.rmtree(download_dir)
        except Exception as e:
            print(f"⚠️ Could not remove download dir: {e}")

    print("🎉 Script finished.")
