import requests
import csv
import time
from datetime import datetime

# Shopify API 설정
SHOPIFY_API_KEY = "fcb0f7def4735642c9f26a307534441e"
SHOPIFY_PASSWORD = "c29ac08e6ed7ddd2d18d38d94875cbe2"
SHOPIFY_STORE_URL = "3abf38-d9.myshopify.com"

# CSV 파일 경로 설정
current_date = datetime.now().strftime('%y%m%d')
csv_file_path = f"C:\\Users\\김남빈\\Desktop\\Coding\\Itemlist_Update\\Active_Items_{current_date}.csv"

# API 헤더 설정
HEADERS = {
    "Content-Type": "application/json",
    "X-Shopify-Access-Token": SHOPIFY_PASSWORD
}

# Shopify API 엔드포인트 설정
BASE_URL = f"https://{SHOPIFY_API_KEY}:{SHOPIFY_PASSWORD}@{SHOPIFY_STORE_URL}/admin/api/2023-01"

def get_product_id_by_barcode(barcode):
    """바코드(Variant Barcode)를 이용하여 Product ID 조회"""
    url = f"{BASE_URL}/variants.json?barcode={barcode}"
    response = requests.get(url, headers=HEADERS)
    
    if response.status_code == 200:
        variants = response.json().get("variants", [])
        if variants:
            return variants[0]["product_id"], variants[0]["id"]
    return None, None

def update_inventory(variant_id, new_stock):
    """Shopify API를 사용하여 상품의 재고(H열) 업데이트"""
    url = f"{BASE_URL}/inventory_levels/set.json"
    data = {
        "location_id": 123456789,  # Shopify 내 매장 Location ID (확인 필요)
        "inventory_item_id": variant_id,
        "available": new_stock
    }
    
    response = requests.post(url, json=data, headers=HEADERS)
    if response.status_code == 200:
        print(f"✅ 재고 업데이트 완료: Variant ID {variant_id} → {new_stock}개")
    else:
        print(f"❌ 재고 업데이트 실패: {response.text}")

def update_variant_price(variant_id, new_price):
    """Shopify API를 사용하여 상품의 할인 가격(K열) 업데이트"""
    url = f"{BASE_URL}/variants/{variant_id}.json"
    data = {
        "variant": {
            "id": variant_id,
            "price": new_price
        }
    }
    
    response = requests.put(url, json=data, headers=HEADERS)
    if response.status_code == 200:
        print(f"✅ 할인 가격 업데이트 완료: Variant ID {variant_id} → {new_price}원")
    else:
        print(f"❌ 할인 가격 업데이트 실패: {response.text}")

def update_product_metafields(product_id, variant_id, expiration_date, store_stock):
    """Shopify API를 사용하여 유통기한(T열) 및 매장별 재고(U,V,W,X열) 업데이트"""
    metafields = [
        {"namespace": "custom", "key": "expiration_date", "value": expiration_date, "type": "string"},
        {"namespace": "custom", "key": "stock_cr", "value": store_stock['CR'], "type": "integer"},
        {"namespace": "custom", "key": "stock_nb", "value": store_stock['NB'], "type": "integer"},
        {"namespace": "custom", "key": "stock_in", "value": store_stock['IN'], "type": "integer"},
        {"namespace": "custom", "key": "stock_br", "value": store_stock['BR'], "type": "integer"},
    ]

    for metafield in metafields:
        url = f"{BASE_URL}/products/{product_id}/metafields.json"
        data = {"metafield": metafield}
        response = requests.post(url, json=data, headers=HEADERS)
        if response.status_code == 200:
            print(f"✅ {metafield['key']} 업데이트 완료: {metafield['value']}")
        else:
            print(f"❌ {metafield['key']} 업데이트 실패: {response.text}")

def process_csv_and_update_shopify():
    """CSV 파일을 읽고 Shopify API를 통해 상품 데이터 업데이트"""
    try:
        with open(csv_file_path, newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)
            
            for row in reader:
                # ✅ 바코드 전처리 (Excel 자동 포맷 제거)
                barcode = row["Variant Barcode"].strip().replace('="', '').replace('"', '')
                
                inventory_quantity = int(row["Variant Inventory Qty"])
                discount_price = float(row["Variant Price"]) if row["Variant Price"] else 0
                expiration_date = row["Expiration_date (product.metafields.custom.expiration_date)"] if row["Expiration_date (product.metafields.custom.expiration_date)"] else ""
                
                store_stock = {
                    "CR": (row["Stock Status CR (product.metafields.custom.stock_cr)"]) if row["Stock Status CR (product.metafields.custom.stock_cr)"] else 0,
                    "NB": (row["Stock Status NB (product.metafields.custom.stock_nb)"]) if row["Stock Status NB (product.metafields.custom.stock_nb)"] else 0,
                    "IN": (row["Stock Status IN (product.metafields.custom.stock_in)"]) if row["Stock Status IN (product.metafields.custom.stock_in)"] else 0,
                    "BR": (row["Stock Status BR (product.metafields.custom.stock_br)"]) if row["Stock Status BR (product.metafields.custom.stock_br)"] else 0,
                }

                product_id, variant_id = get_product_id_by_barcode(barcode)
                
                if variant_id:
                    update_inventory(variant_id, inventory_quantity)
                    update_variant_price(variant_id, discount_price)
                    update_product_metafields(product_id, variant_id, expiration_date, store_stock)
                    
                    time.sleep(1)  # API Rate Limit 방지
                else:
                    print(f"⚠️ 바코드 {barcode}에 해당하는 상품을 찾을 수 없음")

    except FileNotFoundError:
        print(f"❌ CSV 파일을 찾을 수 없습니다: {csv_file_path}")

if __name__ == "__main__":
    process_csv_and_update_shopify()
