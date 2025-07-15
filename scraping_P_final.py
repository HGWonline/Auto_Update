import requests
import json
import sys
from datetime import datetime
import openpyxl
import os

# 슬랙 알림 함수
def send_slack_message(text):
    webhook_url = "https://hooks.slack.com/services/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"
    payload = {"text": text}
    try:
        response = requests.post(webhook_url, json=payload)
        if response.status_code != 200:
            print(f"❗ 슬랙 알림 실패: {response.status_code}")
    except Exception as e:
        print(f"❗ 슬랙 알림 오류: {e}")

def handle_exception_and_exit(context, error):
    message = f"❌ 오류 발생 - {context}\n```\n{str(error)}\n```"
    print(message)
    send_slack_message(message)
    sys.exit(1)

def load_token():
    try:
        with open("token.json", "r") as f:
            token_data = json.load(f)
        return token_data["JSESSIONID"], token_data["X_CSRF_TOKEN"]
    except Exception as e:
        handle_exception_and_exit("token.json 로딩 실패", e)

def fetch_purchase_data(jsessionid, csrf_token):
    url = "https://www.hangawee.com.au/retrievePurchaseList"
    headers = {
        "X-CSRF-TOKEN": csrf_token,
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Origin": "https://www.hangawee.com.au",
        "Referer": "https://www.hangawee.com.au/",
        "User-Agent": "Mozilla/5.0"
    }
    cookies = {
        "JSESSIONID": jsessionid
    }
    payload = {}

    try:
        res = requests.post(url, headers=headers, cookies=cookies, data=payload)
        res.raise_for_status()
        return res.json()
    except Exception as e:
        handle_exception_and_exit("Purchase 데이터 요청 실패", e)

def save_to_excel(purchase_json, filename):
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Purchase"

        ws.append(["Description", "Purchase Date", "Quantity", "Price", "User Name"])

        for item in purchase_json:
            desc = item.get("description", "")
            date = item.get("purchaseDate", "")
            quantity = item.get("quantity", "")
            price = item.get("price", "")
            user = item.get("realName", "")

            ws.append([desc, date, quantity, price, user])

        wb.save(filename)
        print(f"✅ Excel 파일 저장 완료: {filename}")
        send_slack_message(f"✅ Purchase 리스트 크롤링 완료: {filename}")
    except Exception as e:
        handle_exception_and_exit("Excel 저장 실패", e)

def main():
    today = datetime.now().strftime("%y%m%d")
    filename = f"Purchase_{today}.xlsx"

    jsessionid, csrf_token = load_token()
    purchase_data = fetch_purchase_data(jsessionid, csrf_token)
    save_to_excel(purchase_data, filename)

if __name__ == "__main__":
    main()
