import logging
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, FileMessage
import datetime, os, re
import google.auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import utils.gemini_tem as gemi

app = Flask(__name__)

# 設定 LINE API
LINE_CHANNEL_ACCESS_TOKEN = 'LINE_CHANNEL_ACCESS_TOKEN'
LINE_CHANNEL_SECRET = 'LINE_CHANNEL_SECRET'
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# 設定 Google Drive API 和 Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/spreadsheets']
creds, _ = google.auth.load_credentials_from_file('cred.json', scopes=SCOPES)

drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)
spreadsheet_id = 'spreadsheet_id'

# 設定日誌
log_directory = 'logs'
if not os.path.exists(log_directory):
    os.makedirs(log_directory)

log_path = os.path.join(log_directory, 'demo.log')

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levellevelname)s: %(message)s',
                    handlers=[
                        logging.FileHandler(log_path),
                        logging.StreamHandler()
                    ])

@app.route('/')
def home():
    return "Ready vv2!"

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        logging.error("Invalid signature error")
        abort(400)
    return 'OK'

def upload_to_drive(file_path, file_name):
    try:
        file_metadata = {'name': file_name, 'mimeType': 'application/vnd.google-apps.document'}
        media = MediaFileUpload(file_path, resumable=True)
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

        # 取得公開連結
        file_id = file.get('id')
        drive_service.permissions().create(fileId=file_id, body={'type': 'anyone', 'role': 'reader'}).execute()
        link = f"https://drive.google.com/file/d/{file_id}/view"
        logging.info(f"File uploaded successfully: {link}")
        return link
    except Exception as e:
        logging.error(f"Failed to upload to drive: {e}")
        return None

def find_and_update_empty_cell(link, file_name):
    try:
        sheet_range = 'A:C'
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_range
        ).execute()
        
        values = result.get('values', [])
        row_index = len(values) - 1

        while row_index >= 0:
            row = values[row_index]
            if len(row) > 2 and row[2] == '':
                memo_text = row[1]
                logging.info(f"Processing memo text: {memo_text}")
                response = gemi.Gemini_Template(f"""
                【PDF檔案名稱】：{file_name}
                【文字描述】：{memo_text}
                請根據【文字描述】來考慮【PDF檔案名稱】是不是【文字描述】的附檔，基本上【文字描述】就是在講一下公司/銀行的財務狀況, 新聞等等。那 PDF檔案就會是他對應的財務細節等，基本上 PDF檔案名稱就會有那家公司/銀行的名稱，所以公司對應上了，基本就會是對的，但還是要考慮語意，不是有公司名稱就一定是他的附檔

                輸出：請輸出 True 或是 False
                輸出請一定用英文，並且只要輸出 True 或是 False 就好，不需要有任何其他文字在其中
                """)

                logging.info(f"Gemini response: {response}")
                if 'True' in response:
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=f'C{row_index + 1}',
                        valueInputOption='RAW',
                        body={'values': [[link]]}
                    ).execute()
                    logging.info(f"Updated cell C{row_index + 1} with link {link}")
                    return True
            row_index -= 1

        # 如果找不到空的 C 欄位，則新增一筆資料
        timestamp = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        append_to_sheet(date=timestamp, pdf_link=link)
        return False
    except Exception as e:
        logging.error(f"Failed to find and update empty cell: {e}")
        return False

def append_to_sheet(date, text=None, pdf_link=None):
    try:
        sheet_range = 'A:C'
        values = [date, text, pdf_link]
        value_range_body = {
            "values": [values]
        }
        response = sheets_service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=sheet_range,
            valueInputOption="RAW",
            body=value_range_body
        ).execute()
        logging.info(f"Append response: {response}")
    except Exception as e:
        logging.error(f"Failed to append to sheet: {e}")

@handler.add(MessageEvent, message=TextMessage)
def handle_text_message(event):
    text = event.message.text
    user_id = event.source.user_id
    timestamp = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")

    # 檢查是否包含 PDF 連結
    pdf_link_match = re.search(r'http[^\s]+\.pdf', text)
    pdf_link = pdf_link_match.group(0) if pdf_link_match else None

    logging.info(f"Text message received: {text}")
    append_to_sheet(date=timestamp, text=text, pdf_link=pdf_link)

@handler.add(MessageEvent, message=FileMessage)
def handle_file_message(event):
    message_id = event.message.id
    message_content = line_bot_api.get_message_content(message_id)
    ext = event.message.file_name.split('.')[-1]
    file_path = f"tmp/{message_id}.{ext}"

    try:
        with open(file_path, 'wb') as fd:
            for chunk in message_content.iter_content():
                fd.write(chunk)
        logging.info(f"File saved: {file_path}")

        link = upload_to_drive(file_path, event.message.file_name)
        if link:
            find_and_update_empty_cell(link, event.message.file_name)

    except Exception as e:
        logging.error(f"Error processing file message: {e}")
    finally:
        try:
            os.remove(file_path)
            logging.info(f"File removed: {file_path}")
        except Exception as e:
            logging.error(f"Failed to remove file: {e}")

if __name__ == "__main__":
    app.run(threaded=True)
