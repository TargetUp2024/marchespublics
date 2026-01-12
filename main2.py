import os
import time
import re
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

# Google Drive
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# -----------------------------
# CONFIGURATION
# -----------------------------
TARGET_URL = 'https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseDetailsConsultation&refConsultation=968924&orgAcronyme=g3h' 
WEBHOOK_URL = os.getenv("N8N_WEBHOOK_URL1") 

# Google Shared Drive & Service Account
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE", "service_account.json")
SHARED_DRIVE_ID = "1l3fvuCwMpRXMdiJWTXo1-1WASH2NInuQ"  # Your Shared Drive ID
FOLDER_NAME_VAR = "Dossier_Consultation_Ref_968924" 
SCOPES = ['https://www.googleapis.com/auth/drive']

print("🚀 Initializing configuration...")

download_dir = os.path.join(os.getcwd(), "downloads_temp")
extract_dir = os.path.join(os.getcwd(), "extracted_temp")

# Clean start
for d in [download_dir, extract_dir]:
    if os.path.exists(d):
        shutil.rmtree(d)
    os.makedirs(d, exist_ok=True)

# -----------------------------
# SELENIUM SETUP
# -----------------------------
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
# GOOGLE DRIVE FUNCTIONS
# -----------------------------
def get_gdrive_service(sa_key_path):
    """Authenticate and return Google Drive service."""
    if not os.path.exists(sa_key_path):
        print(f"❌ Service Account JSON not found at: {sa_key_path}")
        return None
    try:
        creds = service_account.Credentials.from_service_account_file(
            sa_key_path, scopes=SCOPES
        )
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        print(f"❌ Auth Failed: {e}")
        return None

def create_drive_folder(service, folder_name, parent_id):
    """Create a folder inside Shared Drive or another folder."""
    try:
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id]
        }
        folder = service.files().create(
            body=file_metadata,
            fields='id, webViewLink',
            supportsAllDrives=True
        ).execute()
        print(f"📁 Created Drive Folder: {folder_name} (ID: {folder.get('id')})")
        return folder.get('id'), folder.get('webViewLink')
    except Exception as e:
        print(f"❌ Failed to create folder: {e}")
        return None, None

def upload_file_to_drive(service, file_path, folder_id):
    """Upload a file to a specific folder in Shared Drive."""
    try:
        file_metadata = {
            'name': os.path.basename(file_path),
            'parents': [folder_id]
        }
        media = MediaFileUpload(file_path, resumable=True)
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True
        ).execute()
        print(f"   ☁️ Uploaded: {os.path.basename(file_path)}")
        return file.get('id')
    except Exception as e:
        print(f"   ⚠️ Upload Failed for {file_path}: {e}")
        return None

# -----------------------------
# TEXT EXTRACTION FUNCTIONS
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
folder_drive_link = None

try:
    print(f"\n🔗 Accessing URL: {TARGET_URL}")
    driver.get(TARGET_URL)
    time.sleep(2)

    # Download interaction
    try:
        download_link = wait.until(EC.element_to_be_clickable((By.ID, "ctl0_CONTENU_PAGE_linkDownloadDce")))
        driver.execute_script("arguments[0].scrollIntoView(true);", download_link)
        download_link.click()
        
        # Fill form
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

    # Process file
    downloaded_file = wait_for_download_complete()
    
    if downloaded_file:
        print(f"✅ File downloaded: {os.path.basename(downloaded_file)}")
        file_list_to_process = []

        # Unzip if needed
        if downloaded_file.lower().endswith(".zip"):
            print("📦 Unzipping file locally...")
            if extract_zip(downloaded_file, extract_dir):
                for root, dirs, files in os.walk(extract_dir):
                    for f in files:
                        file_list_to_process.append(os.path.join(root, f))
            else:
                file_list_to_process.append(downloaded_file)
        else:
            file_list_to_process.append(downloaded_file)

        # Setup Drive
        print(f"🚀 Preparing to upload {len(file_list_to_process)} files to Google Drive...")
        drive_service = get_gdrive_service(SERVICE_ACCOUNT_FILE)
        current_folder_id = None

        if drive_service and SHARED_DRIVE_ID:
            current_folder_id, folder_drive_link = create_drive_folder(
                drive_service, FOLDER_NAME_VAR, SHARED_DRIVE_ID
            )

        # Upload & extract text
        extracted_texts = []
        for fpath in file_list_to_process:
            fname = os.path.basename(fpath)

            if drive_service and current_folder_id:
                upload_file_to_drive(drive_service, fpath, current_folder_id)

            ext = os.path.splitext(fname)[1].lower()
            text_chunk = ""
            print(f"   📖 Reading text from: {fname}")
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
            final_output = "Files processed (and uploaded) but no text extracted."

    else:
        final_output = "❌ No file downloaded or timeout occurred."

finally:
    driver.quit()
    if os.path.exists(download_dir): shutil.rmtree(download_dir, ignore_errors=True)
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir, ignore_errors=True)

# -----------------------------
# SEND TO WEBHOOK
# -----------------------------
print("\n" + "="*40)
print("SENDING TO WEBHOOK")
print("="*40)

if WEBHOOK_URL:
    payload = {
        "url": TARGET_URL,
        "status": extraction_status,
        "merged_text": final_output,
        "drive_folder_link": folder_drive_link,
        "timestamp": datetime.now().isoformat()
    }
    
    print(f"📤 Sending data to: {WEBHOOK_URL}")
    try:
        response = requests.post(WEBHOOK_URL, json=payload, timeout=300)
        if response.status_code == 200:
            print("✅ SUCCESS: Data sent to Webhook.")
        else:
            print(f"⚠️ ERROR: Webhook returned status code {response.status_code}")
    except Exception as e:
        print(f"❌ CONNECTION ERROR: {e}")
else:
    print("⚠️ SKIPPED: No WEBHOOK_URL configured.")
    if folder_drive_link:
        print(f"📂 Drive Folder Link: {folder_drive_link}")
