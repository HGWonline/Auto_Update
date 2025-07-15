import pandas as pd
from datetime import datetime

# 현재 날짜 가져오기 (yymmdd 형식)
current_date = datetime.now().strftime('%y%m%d')

# 파일 경로 설정
scraped_file = f'Scraped_{current_date}.xlsx'
itemlist_file = f'Itemlist_{current_date}.xlsx'

# Scraped 파일과 Itemlist 파일을 pandas 데이터프레임으로 읽기
try:
    scraped_df = pd.read_excel(scraped_file)
    itemlist_df = pd.read_excel(itemlist_file)
except Exception as e:
    print(f"Error loading files: {e}")
    exit(1)

# Scraped 데이터를 딕셔너리로 변환 (Scraped 파일에서 Barcode 열을 사용)
if 'Barcode' in scraped_df.columns:
    scraped_df['Barcode'] = scraped_df['Barcode'].astype(str)  # 바코드를 문자열로 변환
    scraped_data = scraped_df.set_index('Barcode').to_dict(orient='index')
else:
    print("Error: 'Barcode' column not found in Scraped file.")
    exit(1)

# Itemlist 바코드 추출 (4개 행씩 묶어서 처리, Booragoon 행에서만 바코드 존재)
def get_itemlist_barcodes(df):
    itemlist_barcodes = {}
    for idx in range(0, len(df), 4):  # 4개 행이 한 세트로 묶여 있음
        barcode_value = df.iloc[idx]['Variant Barcode']  # Booragoon 행에서만 바코드 추출
        if pd.isna(barcode_value) or barcode_value == "":
            continue  # 빈 값은 건너뛰기
        itemlist_barcodes[str(barcode_value)] = idx  # Booragoon 행의 인덱스를 저장
    return itemlist_barcodes

# 바코드 비교
def compare_barcodes(scraped_data, itemlist_barcodes):
    """
    Scraped 데이터와 Itemlist의 바코드를 비교하여 추가할 바코드와 삭제할 바코드를 반환.
    """
    scraped_barcodes = {str(barcode) for barcode in scraped_data.keys()}
    itemlist_barcodes_set = {str(barcode) for barcode in itemlist_barcodes.keys()}
    
    barcodes_to_add = scraped_barcodes - itemlist_barcodes_set
    barcodes_to_delete = itemlist_barcodes_set - scraped_barcodes
    
    return barcodes_to_add, barcodes_to_delete

# 삭제할 바코드 처리 (삭제할 행들만 삭제)
def delete_barcodes(itemlist_barcodes, barcodes_to_delete, df):
    rows_to_delete = sorted([itemlist_barcodes[barcode] for barcode in barcodes_to_delete], reverse=True)
    for row_idx in rows_to_delete:
        df.drop(index=range(row_idx, row_idx + 4), inplace=True)

# 추가할 바코드 처리 (바코드 정보만 추가)
def add_barcodes(barcodes_to_add, df):
    """
    추가할 바코드 정보만 우선적으로 추가하고 나중에 데이터를 입력.
    """
    rows_to_add = []
    for barcode in barcodes_to_add:
        location_values = ['Booragoon', 'Carousel', 'Northbridge', 'Innaloo']
        for location in location_values:
            new_row = {
                'Variant Barcode': barcode if location == 'Booragoon' else '',
                'Store': location,
            }
            rows_to_add.append(new_row)

    df = pd.concat([df, pd.DataFrame(rows_to_add)], ignore_index=True)
    return df

# Scraped 바코드와 Itemlist 바코드를 비교하여 추가 및 삭제할 바코드 확인
itemlist_barcodes = get_itemlist_barcodes(itemlist_df)
barcodes_to_add, barcodes_to_delete = compare_barcodes(scraped_data, itemlist_barcodes)

# 삭제 및 바코드 추가 작업 (데이터는 나중에 입력)
delete_barcodes(itemlist_barcodes, barcodes_to_delete, itemlist_df)
itemlist_df = add_barcodes(barcodes_to_add, itemlist_df)

# 재정렬: Scraped 바코드 순서에 맞춰 Itemlist 재정렬
def reorder_barcodes(scraped_data, df):
    """
    Scraped 파일의 바코드 순서에 맞게 Itemlist 파일의 바코드와 관련된 모든 데이터를 재정렬.
    """
    scraped_barcodes = list(scraped_data.keys())
    itemlist_barcodes = get_itemlist_barcodes(df)

    sorted_data = []
    for barcode in scraped_barcodes:
        barcode_str = str(barcode)  # 문자열로 변환하여 비교
        if barcode_str in itemlist_barcodes:
            start_row = itemlist_barcodes[barcode_str]
            rows_data = df.iloc[start_row:start_row+4].values.tolist()  # 4개 행을 한 세트로 가져옴
            sorted_data.extend(rows_data)
        else:
            print(f"Warning: Barcode {barcode} not found in itemlist_df.")  # 디버깅 메시지

    if not sorted_data:
        raise ValueError("Error: No matching barcodes found for reordering.")

    return pd.DataFrame(sorted_data, columns=df.columns)

# 호출: 바코드를 정렬
itemlist_df = reorder_barcodes(scraped_data, itemlist_df)

# 파일 저장
updated_itemlist_file = f'Updated_Itemlist_{current_date}.xlsx'
itemlist_df.to_excel(updated_itemlist_file, index=False)
print(f"Updated Itemlist saved as '{updated_itemlist_file}'.")
