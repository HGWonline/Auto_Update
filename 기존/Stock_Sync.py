import requests

# Shopify API 설정
SHOP_NAME = '3abf38-d9'  # 사용자 관리자 페이지의 이름
API_VERSION = '2024-07'
ACCESS_TOKEN = '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'

# 스크래핑한 데이터 읽기
def scrape_inventory_data():
    # 이 함수는 Scraping을 통해 재고 정보를 가져오는 함수입니다.
    # 여기에 Scraping 코드가 들어가야 합니다.
    # 예시 데이터
    print("Scraping inventory data...")
    scraped_data = {
        "1234567890123": {
            "booragoon": 10,
            "carousel": 5,
            "innaloo": 20,
            "northbridge": 8
        },
        "9876543210987": {
            "booragoon": 15,
            "carousel": 7,
            "innaloo": 12,
            "northbridge": 10
        }
    }
    print("Scraping complete.")
    return scraped_data

# Shopify 제품 목록 가져오기 (페이지네이션 사용)
def get_shopify_products():
    url = f"https://{SHOP_NAME}.myshopify.com/admin/api/{API_VERSION}/products.json"
    headers = {
        "X-Shopify-Access-Token": ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    products = []
    page_num = 1
    
    while url:
        print(f"Fetching products, page {page_num}...")
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            products += data['products']
            print(f"Page {page_num}: Retrieved {len(data['products'])} products.")
            
            # 페이지네이션을 위한 다음 페이지 링크 추출
            if 'Link' in response.headers:
                link_header = response.headers['Link']
                if 'rel="next"' in link_header:
                    url = link_header.split(';')[0].strip('<>')
                    page_num += 1
                else:
                    url = None  # 더 이상 다음 페이지가 없을 때
            else:
                break  # 페이지네이션 링크가 없을 경우
        else:
            print(f"Error fetching products: {response.status_code}, {response.text}")
            break
    
    print(f"Total products retrieved: {len(products)}")
    return products

# 특정 제품의 재고를 업데이트
def update_inventory_on_shopify(product_id, location_id, available_quantity):
    url = f"https://{SHOP_NAME}.myshopify.com/admin/api/{API_VERSION}/inventory_levels/set.json"
    headers = {
        "X-Shopify-Access-Token": ACCESS_TOKEN,
        "Content-Type": "application/json"
    }
    payload = {
        "location_id": location_id,
        "inventory_item_id": product_id,
        "available": available_quantity
    }
    print(f"Updating inventory for product {product_id} at location {location_id} with quantity {available_quantity}...")
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code == 200:
        print(f"Inventory updated successfully for product {product_id} at location {location_id}.")
    else:
        print(f"Error updating inventory: {response.status_code}, {response.text}")

# 전체 재고 업데이트
def update_inventory():
    print("Starting inventory update process...")
    
    # Scraping 시작
    scraped_data = scrape_inventory_data()
    
    # Shopify에서 제품 목록 가져오기
    products = get_shopify_products()

    # 매장 위치 ID (예시로 추가, 실제 값으로 대체 필요)
    location_ids = {
        "booragoon": "location_id_1",
        "carousel": "location_id_2",
        "innaloo": "location_id_3",
        "northbridge": "location_id_4"
    }

    # 제품들 반복문
    for product in products:
        for variant in product['variants']:
            barcode = variant.get('barcode')
            if barcode in scraped_data:
                inventory_item_id = variant['inventory_item_id']
                inventory_data = scraped_data[barcode]

                # 각 매장의 재고 업데이트
                for location, quantity in inventory_data.items():
                    location_id = location_ids.get(location)
                    if location_id:
                        update_inventory_on_shopify(inventory_item_id, location_id, quantity)

    print("Inventory update process completed.")

# 테스트 실행
update_inventory()
