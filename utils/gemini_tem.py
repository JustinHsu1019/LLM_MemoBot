import requests, json
import sys, os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import utils.config_log as config_log
config, logger, CONFIG_PATH = config_log.setup_config_and_logging()
config.read(CONFIG_PATH)

def Gemini_Template(prompt):
    api_key = config.get("Gemini", 'api_key')
    url = f'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key={api_key}'
    payload = {
        "contents": [
            {
                "parts": [
                    {
                        "text": prompt
                    }
                ]
            }
        ]
    }
    headers = {
        'Content-Type': 'application/json'
    }
    response = (requests.post(url, headers=headers, data=json.dumps(payload))).json()
    return response["candidates"][0]["content"]["parts"][0]["text"]

if __name__ == "__main__":
    prompt = f"""告訴我 CTF 逆向分析的 3 個訣竅，用 json 格式輸出: {{"訣竅1": ,"訣竅2": ,"訣竅3": }}"""
    response = Gemini_Template(prompt)
    print(response)