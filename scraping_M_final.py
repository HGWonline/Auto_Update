import requests
import openpyxl
import pandas as pd
from datetime import datetime
from pathlib import Path
from collections import Counter
import re
import json
import sys
import traceback

### 🛎️ 슬랙 알림 함수
def send_slack_message(text):
    webhook_url = "https://hooks.slack.com/services/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"  # 👈 여기에 Webhook URL 넣기
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

### 🔐 1. token.json 불러오기
try:
    with open("token.json", "r") as f:
        token_data = json.load(f)
        JSESSIONID = token_data["JSESSIONID"]
        X_CSRF_TOKEN = token_data["X_CSRF_TOKEN"]
except Exception as e:
    handle_exception_and_exit("token.json 로딩 실패", e)

cookies = {'JSESSIONID': JSESSIONID}
headers = {
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Origin': 'https://www.hangawee.com.au',
    'Referer': 'https://www.hangawee.com.au/retrieveItemManagementList',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
    'X-CSRF-TOKEN': X_CSRF_TOKEN,
    'X-Requested-With': 'XMLHttpRequest',
}

payload = {
    "start": 0,
    "length": 9999
}

today_str = datetime.now().strftime("%y%m%d")
raw_filename = f"ScrapedM_raw_{today_str}.xlsx"
output_filename = f"ScrapedM_{today_str}.xlsx"

print("📡 서버로부터 데이터 요청 중...")

try:
    response = requests.post(
        "https://www.hangawee.com.au/retrieveItemManagementList",
        headers=headers,
        cookies=cookies,
        json=payload
    )
    if response.status_code != 200:
        raise Exception(f"상태 코드 {response.status_code}")
except Exception as e:
    handle_exception_and_exit("크롤링 요청 실패", e)

try:
    json_data = response.json()
    data = json_data if isinstance(json_data, list) else json_data.get("data", [])
    if not data:
        raise Exception("데이터가 비어 있음")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ItemManagement"
    headers_row = list(data[0].keys())
    ws.append(headers_row)
    for row in data:
        ws.append([row.get(h, "") for h in headers_row])
    wb.save(raw_filename)
    print(f"✅ 크롤링 완료 및 저장: {raw_filename}")
except Exception as e:
    handle_exception_and_exit("JSON 처리 또는 파일 저장 실패", e)

### 🛠 후처리
print("🧹 후처리 시작...")

try:
    df = pd.read_excel(raw_filename)

    df = df[df.iloc[:, 21] != 0]        # retailerInactive != 0
    df = df[df.iloc[:, 1].notna()]      # 상품 이름 not NaN

    def clean_barcode(value):
        if pd.isna(value): return None
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def extract_mode_price(price_string):
        prices = re.findall(r'\d+\.\d+', str(price_string))
        if not prices: return None
        counts = Counter(prices)
        most_common = counts.most_common()
        if len(most_common) == 1 or most_common[0][1] > most_common[1][1]:
            return float(most_common[0][0])
        ca_match = re.search(r'CA:(\d+\.\d+)', price_string)
        return float(ca_match.group(1)) if ca_match else float(most_common[0][0])

    new_rows = []
    for _, row in df.iterrows():
        name = row.iloc[1]
        pak_barcode = clean_barcode(row.iloc[5])
        ea_barcode = clean_barcode(row.iloc[7])
        pak_price_str = str(row.iloc[6])
        ea_price_str = str(row.iloc[8])
        wholesale_price = row.iloc[15]
        quantity_per_box = row.iloc[14]
        ratio = row.iloc[17]

        if pak_barcode and pak_barcode.lower() != "nan":
            pak_wholesale = wholesale_price / quantity_per_box if quantity_per_box else None
            new_rows.append({
                '상품 이름': name,
                'Barcode': pak_barcode,
                'wholesalerPrice': pak_wholesale,
                'retailPrice': extract_mode_price(pak_price_str)
            })

        if ea_barcode and ea_barcode.lower() != "nan":
            ea_wholesale = (wholesale_price / quantity_per_box / ratio) if quantity_per_box and ratio else None
            new_rows.append({
                '상품 이름': name,
                'Barcode': ea_barcode,
                'wholesalerPrice': ea_wholesale,
                'retailPrice': extract_mode_price(ea_price_str)
            })

    final_df = pd.DataFrame(new_rows, columns=['상품 이름', 'Barcode', 'wholesalerPrice', 'retailPrice'])
    final_df = final_df[final_df["Barcode"].notna()]
    final_df = final_df[final_df["Barcode"].str.lower() != "nan"]
    final_df.to_excel(output_filename, index=False)

    msg = f"✅ ScrapedM_{today_str}.xlsx 생성 완료!\n총 {len(final_df)}개 항목 처리됨."
    print(f"🎉 후처리 완료 및 저장: {output_filename}")
    send_slack_message(msg)

except Exception as e:
    handle_exception_and_exit("후처리 중 오류", e)
