import pandas as pd
import os
from datetime import datetime

# 오늘 날짜 (yymmdd 형식)
today_str = datetime.today().strftime('%y%m%d')

# 파일명 설정
input_filename = f"Scraped_{today_str}.xlsx"
output_filename = f"Movement_{today_str}.xlsx"

# 현재 경로
current_dir = os.path.dirname(os.path.abspath(__file__))
input_path = os.path.join(current_dir, input_filename)
output_path = os.path.join(current_dir, output_filename)

# 엑셀 파일 읽기
df = pd.read_excel(input_path)

# 열 이름 추출
columns = df.columns.tolist()
CR_col = columns[2]    # C열
NB_col = columns[5]    # F열
IN_col = columns[8]    # I열
BR_col = columns[11]   # L열
CFC_col = columns[14]  # O열
DESC_col = columns[1]  # B열

# 1차 필터: BR ≤ 5 and (CR ≥ 6 or NB ≥ 6 or IN ≥ 6)
filtered_df = df[
    (df[BR_col] <= 5) & (
        (df[CR_col] >= 6) |
        (df[NB_col] >= 6) |
        (df[IN_col] >= 6)
    )
].copy()

# 2차 필터: CFC_Stock이 'NA', '0.0', 0, 0.0 이외인 행 제거
def is_cfc_valid(value):
    if pd.isna(value):
        return True
    try:
        val = float(value)
        return val == 0.0
    except:
        return str(value).strip().upper() == 'NA'

filtered_df = filtered_df[filtered_df[CFC_col].apply(is_cfc_valid)]

# 3차 필터: Description이 '[SINGLE]'로 끝나는 행 제거
def ends_with_single(desc):
    if not isinstance(desc, str):
        return False
    return desc.strip().upper().endswith('[SINGLE]')

filtered_df = filtered_df[~filtered_df[DESC_col].apply(ends_with_single)]

# 삭제할 열: D, E, G, H, J, K, M, N (index: 3,4,6,7,9,10,12,13)
cols_to_drop = [columns[i] for i in [3,4,6,7,9,10,12,13] if i < len(columns)]
filtered_df.drop(columns=cols_to_drop, inplace=True, errors='ignore')

# 결과 저장
filtered_df.to_excel(output_path, index=False)

print(f"✅ 필터링 완료: '{output_filename}'에 저장되었습니다.")
