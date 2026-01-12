import os
import time
import re
import shutil
import zipfile
import subprocess
import traceback
import unicodedata
import random
import requests  # 👈 Added requests library
from datetime import datetime

# PDF / OCR / DOC
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract
import docx

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import ElementClickInterceptedException

# -----------------------------
# CONFIGURATION
# -----------------------------
# 👇 PUT THE SPECIFIC URL HERE 👇
TARGET_URL = 'https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseDetailsConsultation&refConsultation=968924&orgAcronyme=g3h' 

# 👇 PUT YOUR WEBHOOK URL HERE (or set it as an env variable)
WEBHOOK_URL = os.getenv("N8N_WEBHOOK_URL1") 
# If running locally without env vars, uncomment the line below and paste url:
# WEBHOOK_URL = "https://your-n8n-instance.com/webhook/..."

print("🚀 Initializing configuration...")
download_dir = os.path.join(os.getcwd(), "downloads_temp")
# Clean start
if os.path.exists(download_dir):
    shutil.rmtree(download_dir)
os.makedirs(download_dir, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument("--headless=chrome") # Run headless on servers
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-gpu")

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

service = Service()
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 30)
print("✅ WebDriver initialized.")

PDF_PAGE_LIMIT = 15  # Limit pages per PDF to speed up

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def clean_extracted_text(text):
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\n{2,}", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    cleaned_lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    return "\n".join(cleaned_lines).strip()

def extract_text_from_pdf(file_path):
    text = ""
    try:
        doc = fitz.open(file_path)
        # limit pages
        limit = min(len(doc), PDF_PAGE_LIMIT)
        for i in range(limit):
            text += doc[i].get_text("text") + "\n"
        doc.close()
    except Exception:
        text = ""
    
    # Fallback to OCR if text is empty (scanned PDF)
    if len(text.strip()) < 50:
        try:
            pages = convert_from_path(file_path, last_page=PDF_PAGE_LIMIT)
            for page_image in pages:
                text += pytesseract.image_to_string(page_image, lang="fra+ara+eng") + "\n"
        except Exception as e:
            print(f"⚠️ OCR failed for {file_path}: {e}")
    return clean_extracted_text(text)

def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        return clean_extracted_text(text)
    except Exception:
        return ""

def extract_text_from_doc(file_path):
    try:
        process = subprocess.Popen(["antiword", file_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, _ = process.communicate()
        return clean_extracted_text(stdout.decode("utf-8", errors="ignore"))
    except Exception:
        return ""

def extract_from_zip(file_path):
    try:
        extract_to = os.path.splitext(file_path)[0]
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(extract_to)
        return extract_to
    except Exception as e:
        print(f"⚠️ Failed to unzip: {e}")
        return None

def wait_for_download_complete(timeout=120):
    elapsed = 0
    while elapsed < timeout:
        files = [f for f in os.listdir(download_dir) if not f.endswith(".crdownload") and not f.startswith(".com.google.Chrome")]
        if files:
            # Check if file size is stable
            time.sleep(2) 
            return os.path.join(download_dir, files[0])
        time.sleep(1)
        elapsed += 1
    return None

# -----------------------------
# MAIN LOGIC
# -----------------------------
final_output = ""
extraction_status = "failed"

try:
    print(f"\n🔗 Accessing URL: {TARGET_URL}")
    driver.get(TARGET_URL)
    time.sleep(2)

    # 1. Trigger the download flow (Fill form if necessary)
    try:
        # Click "Télécharger le dossier de consultation"
        download_link = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce")))
        driver.execute_script("arguments[0].scrollIntoView(true);", download_link)
        download_link.click()
        print("📝 Filling access form...")

        # Fill Form Data (Required by the website)
        fields = {
            "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_nom": "Consultant",
            "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_prenom": "External",
            "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_email": "consultant.ext@example.com"
        }

        for fid, value in fields.items():
            inp = wait.until(EC.presence_of_element_located((By.ID, fid)))
            inp.clear()
            inp.send_keys(value)

        # Accept conditions
        checkbox = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions")
        if not checkbox.is_selected():
            checkbox.click()

        # Validate Form
        valider_button = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_validateButton")))
        driver.execute_script("arguments[0].click();", valider_button)

        # Click Final Download Button
        final_button = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload")))
        driver.execute_script("arguments[0].scrollIntoView(true);", final_button)
        final_button.click()
        print("⬇️ Download started...")
        time.sleep(3)

    except Exception as e:
        print(f"❌ Error during download interaction: {e}")

    # 2. Handle File Processing
    downloaded_file = wait_for_download_complete()
    
    if downloaded_file:
        print(f"✅ File downloaded: {os.path.basename(downloaded_file)}")
        file_paths = []
        
        # Check if Zip
        if downloaded_file.lower().endswith(".zip"):
            unzip_dir = extract_from_zip(downloaded_file)
            if unzip_dir:
                for r, _, files in os.walk(unzip_dir):
                    for f in files:
                        file_paths.append(os.path.join(r, f))
        else:
            file_paths.append(downloaded_file)

        extracted_texts = []
        print(f"📂 Processing {len(file_paths)} files inside...")

        for fpath in file_paths:
            fname = os.path.basename(fpath)
            ext = os.path.splitext(fname)[1].lower()
            text_chunk = ""

            print(f"   - Reading: {fname}")
            if ext == ".pdf":
                text_chunk = extract_text_from_pdf(fpath)
            elif ext == ".docx":
                text_chunk = extract_text_from_docx(fpath)
            elif ext == ".doc":
                text_chunk = extract_text_from_doc(fpath)
            
            if text_chunk:
                extracted_texts.append(f"--- START FILE: {fname} ---\n{text_chunk}\n--- END FILE ---\n")

        final_output = "\n".join(extracted_texts)
        if final_output:
            extraction_status = "success"
        else:
            final_output = "Files processed but no text extracted."
    else:
        final_output = "❌ No file downloaded or timeout occurred."

finally:
    # Clean up
    driver.quit()
    if os.path.exists(download_dir):
        shutil.rmtree(download_dir, ignore_errors=True)

# -----------------------------
# FINAL OUTPUT & WEBHOOK
# -----------------------------
print("\n" + "="*40)
print("SENDING TO WEBHOOK")
print("="*40)

if WEBHOOK_URL:
    payload = {
        "url": TARGET_URL,
        "status": extraction_status,
        "merged_text": final_output,
        "timestamp": datetime.now().isoformat()
    }
    
    print(f"📤 Sending data to: {WEBHOOK_URL} (Text len: {len(final_output)})")
    
    try:
        response = requests.post(WEBHOOK_URL, json=payload, timeout=300)
        if response.status_code == 200:
            print("✅ SUCCESS: Data sent to Webhook.")
        else:
            print(f"⚠️ ERROR: Webhook returned status code {response.status_code}")
            print(f"Response: {response.text}")
    except Exception as e:
        print(f"❌ CONNECTION ERROR: Could not send to Webhook. Details: {e}")
else:
    print("⚠️ SKIPPED: No WEBHOOK_URL configured.")
    print("Dumping text locally for review:")
    print(final_output[:2000] + "\n... (truncated)")
