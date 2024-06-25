import logging
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, FileMessage
import datetime, os, re
import google.auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import threading
import queue
from retry import retry
import utils.gemini_tem as gemi
import pytz

app = Flask(__name__)

import utils.config_log as config_log
config, logger, CONFIG_PATH = config_log.setup_config_and_logging()
config.read(CONFIG_PATH)

# 設定 LINE API
LINE_CHANNEL_ACCESS_TOKEN = config.get("line", 'access_token')
LINE_CHANNEL_SECRET = config.get("line", 'secret')
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# 設定 Google Drive API 和 Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/spreadsheets']
creds, _ = google.auth.load_credentials_from_file('cred.json', scopes=SCOPES)

drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)
spreadsheet_id = config.get("line", 'sheet_id')

# 設定日誌
log_directory = 'logs'
if not os.path.exists(log_directory):
    os.makedirs(log_directory)

log_path = os.path.join(log_directory, 'demo.log')

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s',
                    handlers=[
                        logging.FileHandler(log_path),
                        logging.StreamHandler()
                    ])

# 創建任務佇列
task_queue = queue.Queue()

def process_queue():
    while True:
        task = task_queue.get()
        try:
            task()
        finally:
            task_queue.task_done()

# 啟動佇列處理執行緒
threading.Thread(target=process_queue, daemon=True).start()

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

@retry(tries=5, delay=2, backoff=2)
def find_and_update_empty_cell(link, file_name):
    try:
        sheet_range = 'A:C'
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_range
        ).execute()

        values = result.get('values', [])
        row_index = 1  # 開始於第二行

        try:
            gemi_response = gemi.Gemini_Template(f"""
            【PDF檔案名稱】：{file_name}
            
            需求：請幫我從【PDF檔案名稱】提取出你認為這個檔案在講述的"一家"公司名稱

            請注意 1. 提取最短的中文單詞
            
            舉例1: 富邦銀行財務表.pdf
            就提取出"富邦"
            舉例2: morgan 摩根大通相關分析.pdf
            就提取出"摩根"
            
            就是說要能提取出足以辨識這家公司的單詞，但又不要太多字，要是最短的可辨識公司單詞，所以像是如果 舉例1 裡面你提取出"富"就不行，這完全無法從這個字判斷出"富邦銀行"
            
            請注意 2. 注意檔案名稱語意，我需要的只有一家公司名稱，所以要注意這個檔案在講的是哪"一"家公司
            
            舉例: 富邦銀行對摩根大通2020年的財務分析報告.pdf
            就提取出"摩根"
            因為在這個舉例中，富邦銀行只是做出這份分析報告的公司，但這份分析報告描述的是"摩根大通"，而非"富邦銀行"
            """)
        except:
            gemi_response = "GEMINI解析錯誤"

        logging.info(f"Gemini response 公司名稱: {gemi_response}")

        while row_index < len(values):
            row = values[row_index]
            try:
                if row[2] == '':
                    pa = "pass"
            except:
                try:
                    memo_text = row[1]
                except:
                    memo_text = ""
    
                logging.info(f"Processing memo text: {memo_text}")
                if gemi_response in memo_text:
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=f'C{row_index + 1}',
                        valueInputOption='RAW',
                        body={'values': [[link]]}
                    ).execute()
                    logging.info(f"Updated cell C{row_index + 1} with link {link}")
                    return True
            row_index += 1

        # 如果找不到空的 C 欄位，則新增一筆資料在第二行
        timestamp = datetime.datetime.now(pytz.timezone('Asia/Taipei')).strftime("%Y/%m/%d %H:%M:%S")
        append_to_sheet(date=timestamp, pdf_link=link)
        return False
    except Exception as e:
        logging.error(f"Failed to find and update empty cell: {e}")
        raise e

def append_to_sheet(date, text=None, pdf_link=None):
    try:
        sheet_range = 'A:C'
        
        # 先取得目前的值
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_range
        ).execute()
        values = result.get('values', [])
        
        # 插入新的值在第二行
        values.insert(1, [date, text, pdf_link])
        value_range_body = {
            "values": values
        }
        response = sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=sheet_range,
            valueInputOption="RAW",
            body=value_range_body
        ).execute()
        logging.info(f"Append response: {response}")
    except Exception as e:
        logging.error(f"Failed to append to sheet: {e}")
        raise e

@handler.add(MessageEvent, message=TextMessage)
def handle_text_message(event):
    text = event.message.text
    user_id = event.source.user_id
    timestamp = datetime.datetime.now(pytz.timezone('Asia/Taipei')).strftime("%Y/%m/%d %H:%M:%S")

    # 檢查是否包含 PDF 連結
    pdf_link_match = re.search(r'http[^\s]+\.pdf', text)
    pdf_link = pdf_link_match.group(0) if pdf_link_match else None

    logging.info(f"Text message received: {text}")
    task_queue.put(lambda: append_to_sheet(date=timestamp, text=text, pdf_link=pdf_link))

@handler.add(MessageEvent, message=FileMessage)
def handle_file_message(event):
    message_id = event.message.id
    message_content = line_bot_api.get_message_content(message_id)
    ext = event.message.file_name.split('.')[-1]
    file_path = f"tmp/{message_id}.{ext}"

    def save_file():
        try:
            with open(file_path, 'wb') as fd:
                for chunk in message_content.iter_content():
                    fd.write(chunk)
            logging.info(f"File saved: {file_path}")

            task_queue.put(lambda: process_file(file_path, event.message.file_name))
        except Exception as e:
            logging.error(f"Error saving file: {e}")

    save_file()

@retry(tries=5, delay=2, backoff=2)
def process_file(file_path, file_name):
    try:
        link = upload_to_drive(file_path, file_name)
        if link:
            find_and_update_empty_cell(link, file_name)
    except Exception as e:
        logging.error(f"Error processing file: {e}")
    finally:
        try:
            os.remove(file_path)
            logging.info(f"File removed: {file_path}")
        except Exception as e:
            logging.error(f"Failed to remove file: {e}")

if __name__ == "__main__":
    app.run(host="0.0.0.0", threaded=True)
