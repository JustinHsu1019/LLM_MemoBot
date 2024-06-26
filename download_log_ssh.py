import logging
import os
import google.auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from retry import retry

# 設定日誌
log_directory = 'logs'
if not os.path.exists(log_directory):
    os.makedirs(log_directory)

log_path = os.path.join(log_directory, 'upload.log')

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s',
                    handlers=[
                        logging.FileHandler(log_path),
                        logging.StreamHandler()
                    ])

# 設定 Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive.file']
creds, _ = google.auth.load_credentials_from_file('cred.json', scopes=SCOPES)

drive_service = build('drive', 'v3', credentials=creds)

@retry(tries=5, delay=2, backoff=2)
def upload_to_drive(file_path, file_name):
    try:
        file_metadata = {'name': file_name}
        media = MediaFileUpload(file_path, mimetype='application/pdf', resumable=True)
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

        # 取得公開連結
        file_id = file.get('id')
        drive_service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}).execute()
        link = f"https://drive.google.com/file/d/{file_id}/view"
        logging.info(f"File uploaded successfully: {link}")
        return link
    except Exception as e:
        logging.error(f"Failed to upload to drive: {e}")
        raise e

if __name__ == "__main__":
    file_path = input("請輸入檔案路徑: ")
    if not os.path.isfile(file_path):
        logging.error(f"檔案不存在: {file_path}")
    else:
        file_name = os.path.basename(file_path)
        try:
            link = upload_to_drive(file_path, file_name)
            print(f"檔案連結: {link}")
        except Exception as e:
            logging.error(f"無法上傳檔案: {e}")
