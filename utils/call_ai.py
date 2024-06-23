import json

import sys, os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from utils.gpt_tem import GPT_Template
from utils.gemini_tem import Gemini_Template

def call_aied(wait, quest):
    prompt = f"""
請從下列三個選項中選擇出最適合回答 "我的問題" 的答案

"我的問題": {quest}

"選項一": {wait[0]}
"選項二": {wait[1]}
"選項三": {wait[2]}

輸出: 請將選中的那個選項內的 *所有資訊* 都輸出出來

json 格式:
{{
    "輸出": ""
}}
"""
    try:
        res = Gemini_Template(prompt)
        res = json.loads(res)["輸出"]
    except:
        res = "GEMINI解析錯誤"

    return res

if __name__ == "__main__":
    quest = "你喜歡吃什麼?"
    wait = ["我喜歡吃蛋餅", "我喜歡打藍球", "我是個人類"]

    print(call_aied(wait, quest))
