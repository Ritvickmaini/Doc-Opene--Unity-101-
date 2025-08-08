import os
import datetime
import re
import subprocess
import logging

# ========== CONFIG ==========
BASE_PATH = r"\\unityserver\Shared Folders\Public Storage\Weekly Schedules"

try:
    LOG_FILE = os.path.expanduser("~/Desktop/schedule_opener.log")
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
except:
    LOG_FILE = os.path.join(os.getcwd(), "schedule_opener.log")
    
# ========== LOGGING SETUP ==========
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def normalize_date(date_str):
    """
    Normalize date to match filename: remove suffixes, leading zeros, and clean spaces.
    Example: 'Friday 05 February 2025' -> 'Friday 5 February 2025'
    """
    date_str = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', date_str)  # Remove 1st, 2nd, etc.
    date_str = re.sub(r'\b0(\d)', r'\1', date_str)             # Remove leading zeros
    return re.sub(r'\s+', ' ', date_str).strip()               # Clean extra spaces

def open_today_schedule():
    tomorrow = (datetime.datetime.now() + datetime.timedelta(days=1)).strftime("%A %d %B %Y")
    tomorrow_normalized = normalize_date(tomorrow)
    logging.info(f"🔍 Looking for file matching: {tomorrow_normalized}")

    try:
        for folder_name in os.listdir(BASE_PATH):
            logging.info(f"📁 Checking folder: {folder_name}")

            folder_path = os.path.join(BASE_PATH, folder_name)
            if not os.path.isdir(folder_path):
                continue

            if not folder_name.lower().startswith("schedules ending"):
                continue

            logging.info(f"🔎 Entering: {folder_path}")

            for file_name in os.listdir(folder_path):
                if not file_name.lower().endswith((".docx", ".pdf")):
                    continue

                file_name_no_ext = os.path.splitext(file_name)[0]
                normalized_file_name = normalize_date(file_name_no_ext)

                logging.info(f"📄 Found file: {file_name}")
                logging.info(f"👉 Normalized filename: {normalized_file_name}")

                if normalized_file_name == tomorrow_normalized:
                    full_file_path = os.path.join(folder_path, file_name)
                    logging.info(f"✅ MATCH FOUND: {full_file_path}")

                    # Open file with default Windows application
                    os.startfile(full_file_path)
                    return

        logging.warning("⚠️ No matching file found for today's date.")

    except Exception as e:
        logging.error(f"❌ Error occurred: {str(e)}")

if __name__ == "__main__":
    open_today_schedule()
