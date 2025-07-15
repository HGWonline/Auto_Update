import openpyxl
import csv
import re
from datetime import datetime

# 현재 날짜 가져오기 (yymmdd 형식)
current_date = datetime.now().strftime('%y%m%d')

# 파일 경로 설정
scraped_file = f'Scraped_{current_date}.xlsx'
csv_file = f'Stock_{current_date}.csv'

# 특수 문자 제거 함수
def clean_string(s):
    # 공백을 '-'로 대체, 특수문자는 전부 제거
    s = re.sub(r'[^\w\s]', '', s)  # 특수 문자 제거
    s = s.replace(' ', '-')  # 공백을 '-'로 대체
    return s

# Scraped_yymmdd.xlsx 파일 로드
print(f"Loading scraped file: {scraped_file}")
try:
    wb = openpyxl.load_workbook(scraped_file, read_only=True)
    sheet = wb.active
    print("Scraped file loaded successfully.")
except Exception as e:
    print(f"Error loading scraped file: {e}")
    sheet = None

# CSV 파일 생성 및 데이터 저장
with open(csv_file, mode='w', newline='', encoding='utf-8-sig') as file:
    writer = csv.writer(file)
    
    # CSV 파일의 헤더 작성
    writer.writerow(['Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Option3 Value', 'SKU', 'Location', 'On Hand'])

    # Scraped 파일에서 데이터 읽기
    if sheet:
        print("Reading scraped data and writing to CSV...")
        try:
            # 각 행에 대해 데이터 추출
            for row in sheet.iter_rows(min_row=2, values_only=True):
                barcode = row[0]  # BarCode
                handle = row[1]  # Handle
                description = row[2]  # Description
                carousel_stock = row[3] if row[3] is not None else 0  # CAROUSEL Stock
                northbridge_stock = row[4] if row[4] is not None else 0  # NORTHBRIDGE Stock
                innaloo_stock = row[5] if row[5] is not None else 0  # INNALOO Stock
                booragoon_stock = row[6] if row[6] is not None else 0  # BOORAGOON Stock
                
                # Handle과 Title 구성 (Handle은 파일에서 불러옴)
                handle = clean_string(handle)
                title = description
                
                # Option1 Value, Option2 Value, Option3 Value 설정
                option1_value = 'Default Title'
                option2_value = ''
                option3_value = ''
                
                # SKU에 BarCode 입력
                sku = str(barcode)
                
                # 각 Location별로 데이터 생성 (총 4개 행으로 나누어 작성)
                locations = ['HanGaWee_Carousel', 'HanGaWee_Northbridge', 'HanGaWee_Innaloo', 'HanGaWee_Booragoon']
                stocks = [carousel_stock, northbridge_stock, innaloo_stock, booragoon_stock]
                
                for location, stock in zip(locations, stocks):
                    writer.writerow([handle, title, option1_value, option2_value, option3_value, sku, location, stock])
            
            print(f"Data successfully written to {csv_file}")

        except Exception as e:
            print(f"Error processing data: {e}")
