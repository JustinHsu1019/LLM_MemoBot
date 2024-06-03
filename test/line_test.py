from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, FileMessage, TextSendMessage
import os, datetime

app = Flask(__name__)

LINE_CHANNEL_ACCESS_TOKEN = 'LINE_CHANNEL_ACCESS_TOKEN'
LINE_CHANNEL_SECRET = 'LINE_CHANNEL_SECRET'
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

@app.route('/')
def home():
    return "Ready!"

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_text_message(event):
    text = event.message.text
    if "memo" in text.lower():
        user_id = event.source.user_id
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        memo_content = f"{user_id}: {text}\n"

        with open("/home/certidtest/mysite/memo.txt", "a") as f:
            f.write(memo_content)

        memo_file_path = f"/home/certidtest/mysite/memos/{user_id}_{timestamp}.txt"
        with open(memo_file_path, "w") as f:
            f.write(memo_content)

        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="Memo saved!"))

@handler.add(MessageEvent, message=FileMessage)
def handle_file_message(event):
    message_id = event.message.id
    message_content = line_bot_api.get_message_content(message_id)
    ext = event.message.file_name.split('.')[-1]
    file_path = f"/home/certidtest/mysite/downloads/{message_id}.{ext}"

    with open(file_path, 'wb') as fd:
        for chunk in message_content.iter_content():
            fd.write(chunk)

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text="File received!"))

if __name__ == "__main__":
    try:
        os.makedirs("/home/certidtest/mysite/downloads", exist_ok=True)
        os.makedirs("/home/certidtest/mysite/memos", exist_ok=True)
    except Exception as e:
        print(f"Error creating directory: {e}")
    app.run(threaded=True)
