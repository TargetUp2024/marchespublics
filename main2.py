import os
import time
import shutil
import zipfile
import subprocess
import unicodedata
import requests
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

# -----------------------------
# CONFIGURATION
# -----------------------------
TARGET_URL = 'https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseDetailsConsultation&refConsultation=968924&orgAcronyme=g3h' 
WEBHOOK_URL = os.getenv("N8N_WEBHOOK_URL1") 

print("🚀 Initializing configuration...")
download_dir = os.path.join(os.getcwd(), "downloads_temp")
extract_dir = os.path.join(os.getcwd(), "extracted_temp")

# Clean start
for d in [download_dir, extract_dir]:
    if os.path.exists(d):
        shutil.rmtree(d)
    os.makedirs(d, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument("--headless=chrome") 
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

PDF_PAGE_LIMIT = 15 

# -----------------------------
# TEXT EXTRACTION HELPER FUNCTIONS
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
        limit = min(len(doc), PDF_PAGE_LIMIT)
        for i in range(limit):
            text += doc[i].get_text("text") + "\n"
        doc.close()
    except Exception:
        text = ""
    
    if len(text.strip()) < 50:
        try:
            pages = convert_from_path(file_path, last_page=PDF_PAGE_LIMIT)
            for page_image in pages:
                text += pytesseract.image_to_string(page_image, lang="fra+ara+eng") + "\n"
        except Exception:
            pass
    return clean_extracted_text(text)

def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return clean_extracted_text("\n".join(p.text for p in doc.paragraphs if p.text.strip()))
    except Exception:
        return ""

def extract_text_from_doc(file_path):
    try:
        process = subprocess.Popen(["antiword", file_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, _ = process.communicate()
        return clean_extracted_text(stdout.decode("utf-8", errors="ignore"))
    except Exception:
        return ""

def extract_zip(zip_path, extract_to_folder):
    try:
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(extract_to_folder)
        return True
    except Exception as e:
        print(f"⚠️ Failed to unzip: {e}")
        return False

def wait_for_download_complete(timeout=120):
    elapsed = 0
    while elapsed < timeout:
        files = [f for f in os.listdir(download_dir) if not f.endswith(".crdownload") and not f.startswith(".com.google.Chrome")]
        if files:
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
downloaded_file_path = None

try:
    print(f"\n🔗 Accessing URL: {TARGET_URL}")
    driver.get(TARGET_URL)
    time.sleep(2)

    # 1. Download Interaction
    try:
        download_link = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce")))
        driver.execute_script("arguments[0].scrollIntoView(true);", download_link)
        download_link.click()
        
        # Fill Form
        fields = {
            "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_nom": "Consultant",
            "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_prenom": "External",
            "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_email": "consultant.ext@example.com"
        }
        for fid, value in fields.items():
            wait.until(EC.presence_of_element_located((By.ID, fid))).send_keys(value)

        checkbox = driver.find_element(By.ID, "ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions")
        if not checkbox.is_selected(): checkbox.click()

        valider = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_validateButton")))
        driver.execute_script("arguments[0].click();", valider)

        final_dl = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload")))
        driver.execute_script("arguments[0].scrollIntoView(true);", final_dl)
        final_dl.click()
        print("⬇️ Download started...")
        time.sleep(3)

    except Exception as e:
        print(f"❌ Error during download interaction: {e}")

    # 2. Process File
    downloaded_file_path = wait_for_download_complete()
    
    if downloaded_file_path:
        print(f"✅ File downloaded: {os.path.basename(downloaded_file_path)}")
        
        # We process files locally for TEXT extraction, but we upload the Original Zip later
        file_list_to_read_text = []

        # A. Unzip Locally (Only for extracting text)
        if downloaded_file_path.lower().endswith(".zip"):
            print("📦 Unzipping file locally for text extraction...")
            if extract_zip(downloaded_file_path, extract_dir):
                for root, dirs, files in os.walk(extract_dir):
                    for f in files:
                        file_list_to_read_text.append(os.path.join(root, f))
            else:
                file_list_to_read_text.append(downloaded_file_path)
        else:
            file_list_to_read_text.append(downloaded_file_path)

        # B. Extract Text
        extracted_texts = []
        for fpath in file_list_to_read_text:
            fname = os.path.basename(fpath)
            ext = os.path.splitext(fname)[1].lower()
            text_chunk = ""
            
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
    driver.quit()
    # We do NOT delete download_dir yet because we need to send the file

# -----------------------------
# FINAL OUTPUT & WEBHOOK
# -----------------------------
print("\n" + "="*40)
print("SENDING TO WEBHOOK")
print("="*40)

if WEBHOOK_URL:
    
    # 1. JSON Data
    payload_data = {
        "url": TARGET_URL,
        "status": extraction_status,
        "merged_text": final_output,
        "timestamp": datetime.now().isoformat()
    }

    # 2. File Data (The ZIP file)
    files_payload = {}
    if downloaded_file_path and os.path.exists(downloaded_file_path):
        print(f"📎 Attaching file: {os.path.basename(downloaded_file_path)}")
        # Open file in Binary mode
        files_payload['file'] = (
            os.path.basename(downloaded_file_path), 
            open(downloaded_file_path, 'rb'), 
            'application/zip'
        )
    
    print(f"📤 Sending data to: {WEBHOOK_URL}")
    
    try:
        # Sending multipart/form-data
        response = requests.post(WEBHOOK_URL, data=payload_data, files=files_payload, timeout=300)
        
        if response.status_code == 200:
            print("✅ SUCCESS: ZIP File and Text sent to Webhook.")
        elif response.status_code == 404:
            print("❌ ERROR 404: Webhook URL not found. Check if workflow is Active in N8N.")
        else:
            print(f"⚠️ ERROR: Webhook returned status code {response.status_code}")
            print(f"Response: {response.text}")

    except Exception as e:
        print(f"❌ CONNECTION ERROR: {e}")
        
    # Close file if it was opened
    if 'file' in files_payload:
        files_payload['file'][1].close()

else:
    print("⚠️ SKIPPED: No WEBHOOK_URL configured.")

# Final Cleanup
if os.path.exists(download_dir): shutil.rmtree(download_dir, ignore_errors=True)
if os.path.exists(extract_dir): shutil.rmtree(extract_dir, ignore_errors=True)
