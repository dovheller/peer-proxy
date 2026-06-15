"""
הורדת תיקיית Google Drive לדיסק המקומי
תיקייה: עו"ד אוסדיטשר / הלר גפנר
יעד: D:\12\עוד אוסדיטשר - הלר גפנר\

דרישות:
    pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client

שימוש:
    python download_drive_folder.py

הסקריפט ייצור קובץ לוג: download_log.txt
"""

import os
import sys
import time
import logging
import re
import io
from pathlib import Path
from datetime import datetime

try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from googleapiclient.errors import HttpError
except ImportError:
    print("ERROR: חסרות תלויות. הרץ:")
    print("  pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client")
    sys.exit(1)

# ─── הגדרות ─────────────────────────────────────────────────────────────────

FOLDER_ID   = "1sS1o-KNRcWQRu12gs_Su9-i2wnKxqf4R"
DEST_ROOT   = Path(r"D:\12")
LOG_FILE    = Path(r"D:\12\download_log.txt")
TOKEN_FILE  = "token.json"
CREDS_FILE  = "credentials.json"   # יש להוריד מ-Google Cloud Console

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

# פורמטי ייצוא עבור קבצי Google Docs
EXPORT_FORMATS = {
    "application/vnd.google-apps.document":
        ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"),
    "application/vnd.google-apps.spreadsheet":
        ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    "application/vnd.google-apps.presentation":
        ("application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx"),
    "application/vnd.google-apps.drawing":
        ("image/png", ".png"),
    "application/vnd.google-apps.script":
        ("application/vnd.google-apps.script+json", ".json"),
    "application/vnd.google-apps.form":
        ("application/pdf", ".pdf"),
}

MAX_RETRIES  = 5
RETRY_DELAY  = 3   # שניות בין ניסיונות


# ─── לוגר ────────────────────────────────────────────────────────────────────

def setup_logger():
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    fmt = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger(__name__)


log = setup_logger()


# ─── אימות ───────────────────────────────────────────────────────────────────

def authenticate() -> Credentials:
    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            log.info("מרענן טוקן...")
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDS_FILE):
                log.error(
                    f"קובץ {CREDS_FILE} לא נמצא.\n"
                    "1. היכנס ל-https://console.cloud.google.com\n"
                    "2. צור פרויקט > הפעל Google Drive API\n"
                    "3. צור OAuth 2.0 Client ID (Desktop)\n"
                    "4. הורד את ה-JSON ושמור אותו כ-credentials.json\n"
                    "   באותה תיקייה שבה נמצא הסקריפט"
                )
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(creds.to_json())
        log.info("טוקן נשמר ב-token.json")

    return creds


# ─── כלי עזר ─────────────────────────────────────────────────────────────────

def sanitize_name(name: str) -> str:
    """הסרת תווים לא-חוקיים לשמות קבצים ב-Windows."""
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = name.strip(". ")
    return name or "_unnamed_"


def list_folder_contents(service, folder_id: str) -> list:
    """מחזיר רשימה של כל הפריטים בתיקייה (כולל דפים מרובים)."""
    items = []
    page_token = None
    query = f"'{folder_id}' in parents and trashed = false"

    while True:
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = service.files().list(
                    q=query,
                    spaces="drive",
                    fields="nextPageToken, files(id, name, mimeType, size)",
                    pageToken=page_token,
                    pageSize=1000,
                ).execute()
                break
            except HttpError as e:
                if attempt == MAX_RETRIES:
                    log.error(f"כשל בקריאת תיקייה {folder_id}: {e}")
                    return items
                log.warning(f"ניסיון {attempt}/{MAX_RETRIES} נכשל, ממתין {RETRY_DELAY}s...")
                time.sleep(RETRY_DELAY * attempt)

        items.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    return items


def download_file(service, file_id: str, mime_type: str, dest_path: Path) -> bool:
    """מוריד קובץ בודד. מחזיר True בהצלחה."""
    dest_path.parent.mkdir(parents=True, exist_ok=True)

    if dest_path.exists():
        log.info(f"  [דלג] קיים: {dest_path.name}")
        return True

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            if mime_type in EXPORT_FORMATS:
                export_mime, ext = EXPORT_FORMATS[mime_type]
                if not dest_path.suffix:
                    dest_path = dest_path.with_suffix(ext)
                request = service.files().export_media(fileId=file_id, mimeType=export_mime)
            else:
                request = service.files().get_media(fileId=file_id)

            buf = io.BytesIO()
            downloader = MediaIoBaseDownload(buf, request, chunksize=10 * 1024 * 1024)
            done = False
            while not done:
                _, done = downloader.next_chunk()

            dest_path.parent.mkdir(parents=True, exist_ok=True)
            dest_path.write_bytes(buf.getvalue())
            log.info(f"  [הורד] {dest_path.name}  ({len(buf.getvalue()):,} bytes)")
            return True

        except HttpError as e:
            if e.resp.status == 403 and "exportSizeLimitExceeded" in str(e):
                log.warning(f"  [גדול] {dest_path.name} — מנסה PDF...")
                try:
                    request = service.files().export_media(fileId=file_id, mimeType="application/pdf")
                    buf = io.BytesIO()
                    downloader = MediaIoBaseDownload(buf, request, chunksize=10 * 1024 * 1024)
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()
                    pdf_path = dest_path.with_suffix(".pdf")
                    pdf_path.write_bytes(buf.getvalue())
                    log.info(f"  [הורד PDF] {pdf_path.name}")
                    return True
                except Exception as e2:
                    log.error(f"  [שגיאה] {dest_path.name}: {e2}")
                    return False
            elif attempt == MAX_RETRIES:
                log.error(f"  [כשל] {dest_path.name}: {e}")
                return False
            else:
                log.warning(f"  ניסיון {attempt}/{MAX_RETRIES} נכשל ({e}), ממתין...")
                time.sleep(RETRY_DELAY * attempt)

        except Exception as e:
            if attempt == MAX_RETRIES:
                log.error(f"  [כשל] {dest_path.name}: {e}")
                return False
            log.warning(f"  ניסיון {attempt}/{MAX_RETRIES} נכשל ({e}), ממתין...")
            time.sleep(RETRY_DELAY * attempt)

    return False


# ─── לוגיקה רקורסיבית ────────────────────────────────────────────────────────

stats = {"downloaded": 0, "skipped": 0, "failed": 0, "folders": 0}


def process_folder(service, folder_id: str, local_path: Path, depth: int = 0):
    indent = "  " * depth
    local_path.mkdir(parents=True, exist_ok=True)
    items = list_folder_contents(service, folder_id)

    if not items:
        log.info(f"{indent}[ריקה] {local_path.name}")
        return

    log.info(f"{indent}[תיקייה] {local_path.name}  ({len(items)} פריטים)")
    stats["folders"] += 1

    # תיקיות קודם (BFS יתכן, אך DFS פשוט יותר)
    for item in items:
        name      = sanitize_name(item["name"])
        mime      = item["mimeType"]
        item_id   = item["id"]

        if mime == "application/vnd.google-apps.folder":
            process_folder(service, item_id, local_path / name, depth + 1)
        else:
            # קבוצת Google Docs — ייצוא; אחרת — הורדה ישירה
            if mime in EXPORT_FORMATS:
                _, ext = EXPORT_FORMATS[mime]
                dest = local_path / (name + ext)
            else:
                dest = local_path / name

            ok = download_file(service, item_id, mime, dest)
            if ok:
                if dest.exists():
                    stats["downloaded"] += 1
                else:
                    stats["skipped"] += 1
            else:
                stats["failed"] += 1


# ─── נקודת כניסה ─────────────────────────────────────────────────────────────

def main():
    log.info("=" * 60)
    log.info(f"התחלת הורדה: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log.info(f"תיקיית יעד: {DEST_ROOT}")
    log.info("=" * 60)

    creds   = authenticate()
    service = build("drive", "v3", credentials=creds)

    process_folder(service, FOLDER_ID, DEST_ROOT)

    log.info("=" * 60)
    log.info("סיכום:")
    log.info(f"  תיקיות:    {stats['folders']}")
    log.info(f"  הורדו:     {stats['downloaded']}")
    log.info(f"  דולגו:     {stats['skipped']}  (כבר קיימים)")
    log.info(f"  נכשלו:     {stats['failed']}")
    log.info(f"  לוג מלא:   {LOG_FILE}")
    log.info("=" * 60)

    if stats["failed"]:
        log.warning(f"שים לב: {stats['failed']} קבצים לא הורדו. ראה לוג לפרטים.")
        sys.exit(1)


if __name__ == "__main__":
    main()
