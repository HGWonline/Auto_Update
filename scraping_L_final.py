import requests
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
import json
import re
import sys
from collections import Counter

# 🛎️ 슬랙 알림 함수
def send_slack_message(text):
    webhook_url = "https://hooks.slack.com/services/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"  # 👈 Webhook URL 직접 지정
    payload = {"text": text}
    try:
        response = requests.post(webhook_url, json=payload)
        if response.status_code != 200:
            print(f"❗ 슬랙 알림 실패: {response.status_code}")
    except Exception as e:
        print(f"❗ 슬랙 알림 오류: {e}")

# ❌ 예외 발생 시 슬랙으로 전송 후 종료
def handle_exception_and_exit(context, error):
    message = f"❌ 오류 발생 - {context}\n```\n{str(error)}\n```"
    print(message)
    send_slack_message(message)
    sys.exit(1)

# token.json에서 인증 정보 로드
try:
    with open("token.json", "r") as f:
        token_data = json.load(f)
        JSESSIONID = token_data["JSESSIONID"]
        X_CSRF_TOKEN = token_data["X_CSRF_TOKEN"]
except Exception as e:
    send_slack_message(f"❌ token.json 로딩 실패: {e}")
    sys.exit(1)

# 요청 설정
cookies = {'JSESSIONID': JSESSIONID}
headers = {
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Origin': 'https://www.hangawee.com.au',
    'Referer': 'https://www.hangawee.com.au/?page=mitems',
    'User-Agent': 'Mozilla/5.0',
    'X-CSRF-TOKEN': X_CSRF_TOKEN,
    'X-Requested-With': 'XMLHttpRequest',
}
url = "https://www.hangawee.com.au/retrieveItemList"

# API 요청
try:
    response = requests.post(url, headers=headers, cookies=cookies)
    data = response.json()
except Exception as e:
    send_slack_message(f"❌ 요청 실패 또는 JSON 파싱 실패: {e}")
    sys.exit(1)

# 매장 정보
stores = [
    ("carouselStockReport", "CR"),
    ("northBridgeStockReport", "NB"),
    ("innalooStockReport", "IN"),
    ("booragoonStockReport", "BR")
]

# 각 보고서 텍스트 파싱
def parse_store_report(report_text):
    if not report_text or report_text == "NotExist":
        return [], None

    soup = BeautifulSoup(report_text.replace('<BR>', '<br>'), "html.parser")
    text = soup.get_text(separator=' ').strip()

    # 할인 정보 추출
    discount_match = re.search(
        r'\[(\d{2}/\d{2}/\d{4})~\s*(\d{2}/\d{2}/\d{4})]:\s*([\d.]+)%\s*\$(\d+\.\d+)', text
    )
    discount_info = None
    if discount_match:
        _, discount_end, discount_rate, discount_price = discount_match.groups()
        discount_info = {
            "discount_end": discount_end,
            "discount_rate": float(discount_rate),
            "discount_price": float(discount_price)
        }

    # 유통기한별 수량
    expiration_blocks = re.findall(r'(\d{2}/\d{2}/\d{4})\[(\-?\d+\.?\d*)\]', text)
    if expiration_blocks:
        blocks = [(exp, float(qty)) for exp, qty in expiration_blocks if float(qty) > 0]
        return blocks, discount_info

    # 유통기한 정보 없고 그냥 숫자만 있는 경우
    match = re.match(r'^(\d+)', text)
    if match:
        return [("", float(match.group(1)))], None

    return [], discount_info

# 전체 데이터 가공
rows = []
for item in data:
    barcode = item.get("barCode")
    desc = item.get("purchaseDescription")
    if not barcode or not desc:
        continue

    discount_candidates = []
    store_blocks = {}

    for key, label in stores:
        report = item.get(key, "")
        blocks, discount_info = parse_store_report(report)
        store_blocks[label] = {"blocks": blocks, "discount": discount_info}
        if discount_info:
            discount_candidates.append((
                discount_info["discount_rate"],
                discount_info["discount_end"],
                discount_info["discount_price"]
            ))

    # 할인 기준 결정
    if discount_candidates:
        max_rate = max(d[0] for d in discount_candidates)
        filtered = [d for d in discount_candidates if d[0] == max_rate]
        sorted_filtered = sorted(filtered, key=lambda x: datetime.strptime(x[1], "%d/%m/%Y"), reverse=True)
        longest_end = sorted_filtered[0][1]
        final_price = sorted_filtered[0][2]
        cutoff_date = datetime.strptime(longest_end, "%d/%m/%Y")
    else:
        common_end = ""
        final_price = ""
        cutoff_date = None

    def calc_store_data(label):
        blocks = store_blocks[label]["blocks"]
        if not blocks:
            return 0, "", ""

        if cutoff_date:
            filtered = [
                (exp, qty) for exp, qty in blocks
                if not exp or (exp and datetime.strptime(exp, "%d/%m/%Y") <= cutoff_date)
            ]
        else:
            filtered = blocks

        if not filtered:
            return 0, "", ""

        total_stock = int(sum(qty for _, qty in filtered))

        # 유통기한이 유효한 경우 중 가장 이른 날짜 찾기
        dated_blocks = [x for x in filtered if x[0]]
        if dated_blocks:
            earliest_exp = min(
                dated_blocks,
                key=lambda x: datetime.strptime(x[0], "%d/%m/%Y")
            )[0]
        else:
            earliest_exp = ""

        return total_stock, final_price if final_price else "", earliest_exp

    cr_stock, cr_price, cr_exp = calc_store_data("CR")
    nb_stock, nb_price, nb_exp = calc_store_data("NB")
    in_stock, in_price, in_exp = calc_store_data("IN")
    br_stock, br_price, br_exp = calc_store_data("BR")

    # CFC 재고
    cfc = item.get("cfcStockReport", "NotExist")
    if cfc == "NotExist":
        cfc_stock = "NA"
    else:
        m = re.search(r"(\d+(\.\d+)?)", cfc)
        cfc_stock = m.group(1) if m else "NA"

    rows.append([
        barcode, desc,
        cr_stock, cr_price, cr_exp,
        nb_stock, nb_price, nb_exp,
        in_stock, in_price, in_exp,
        br_stock, br_price, br_exp,
        cfc_stock
    ])

# 저장
df = pd.DataFrame(rows, columns=[
    "Barcode", "Description",
    "CR_Stock", "CR_Discount", "CR_Expiration",
    "NB_Stock", "NB_Discount", "NB_Expiration",
    "IN_Stock", "IN_Discount", "IN_Expiration",
    "BR_Stock", "BR_Discount", "BR_Expiration",
    "CFC_Stock"
])

file_name = f"Scraped_{datetime.now().strftime('%y%m%d')}.xlsx"
df.to_excel(file_name, index=False)

send_slack_message(f"✅ {file_name} 생성 완료!")
