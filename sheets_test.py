import logging
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, FileMessage, TextSendMessage
import datetime, os
import google.auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

app = Flask(__name__)

# 設定 LINE API
LINE_CHANNEL_ACCESS_TOKEN = 'LINE_CHANNEL_ACCESS_TOKEN'
LINE_CHANNEL_SECRET = 'LINE_CHANNEL_SECRET'
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# 設定 Google Drive API 和 Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/spreadsheets']
creds, _ = google.auth.load_credentials_from_file('/home/certidtest/mysite/credentials.json', scopes=SCOPES)

drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)
spreadsheet_id = 'spreadsheet_id'

# 設定日誌
log_directory = '/home/certidtest/mysite/logs'
if not os.path.exists(log_directory):
    os.makedirs(log_directory)

log_path = os.path.join(log_directory, 'demo.log')

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(levelname)s: %(message)s',
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

def append_to_sheet(link):
    try:
        sheet_range = 'A:A'
        value_range_body = {
            "values": [[link]]
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
    if "memo" in text.lower():
        user_id = event.source.user_id
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        memo_content = f"{user_id}: {text}\n"

        memo_file_path = f"/home/certidtest/mysite/tmp/{user_id}_{timestamp}.txt"
        with open(memo_file_path, "w") as f:
            f.write(memo_content)

        link = upload_to_drive(memo_file_path, f"{user_id}_{timestamp}.txt")
        if link:
            append_to_sheet(link)
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="Memo 已儲存並上傳至雲端！"))
        else:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="Memo 儲存失敗，請稍後再試！"))

        try:
            os.remove(memo_file_path)
        except Exception as e:
            logging.error(f"Failed to remove file: {e}")

@handler.add(MessageEvent, message=FileMessage)
def handle_file_message(event):
    message_id = event.message.id
    message_content = line_bot_api.get_message_content(message_id)
    ext = event.message.file_name.split('.')[-1]
    file_path = f"/home/certidtest/mysite/tmp/{message_id}.{ext}"

    with open(file_path, 'wb') as fd:
        for chunk in message_content.iter_content():
            fd.write(chunk)

    link = upload_to_drive(file_path, event.message.file_name)
    if link:
        append_to_sheet(link)
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="檔案已收到並上傳至雲端！"))
    else:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="檔案上傳失敗，請稍後再試！"))

    try:
        os.remove(file_path)
    except Exception as e:
        logging.error(f"Failed to remove file: {e}")

if __name__ == "__main__":
    app.run(threaded=True)
