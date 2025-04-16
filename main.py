import json
import gspread
import openai
from oauth2client.service_account import ServiceAccountCredentials
from utils import process_review_file
import os

def get_google_client():
    creds_dict = json.loads(os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON"))
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(creds)

def main():
    gc = get_google_client()
    openai_key = os.getenv("OPENAI_API_KEY")

    # 管理用マスターシート（1列目にURLあり）
    master_url = os.getenv("MASTER_SPREADSHEET_URL")
    master_sheet = gc.open_by_url(master_url).sheet1

    urls = master_sheet.col_values(1)[1:]  # A列の2行目以降がPoC対象

    for url in urls:
        try:
            spreadsheet = gc.open_by_url(url)
            print(f"Processing: {url}")
            process_review_file(spreadsheet, openai_key)
        except Exception as e:
            print(f"Error processing {url}: {e}")

if __name__ == "__main__":
    main()
