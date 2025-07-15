import os
import datetime
import pandas as pd

def create_combined_excel():
    today_str = datetime.datetime.today().strftime('%y%m%d')

    base_dir = os.path.dirname(os.path.abspath(__file__))
    scraped_filename = f"Scraped_{today_str}.xlsx"
    itemlist_filename = f"Itemlist_{today_str}.xlsx"

    scraped_path = os.path.join(base_dir, scraped_filename)
    itemlist_path = os.path.join(base_dir, itemlist_filename)

    scraped_df = pd.read_excel(scraped_path)
    itemlist_df = pd.read_excel(itemlist_path)

    df = pd.DataFrame({
        'Barcode': scraped_df.iloc[:, 0],
        'Description': scraped_df.iloc[:, 1],
        'Stock': scraped_df.iloc[:, 11],
        'Discount': itemlist_df.iloc[:, 10],
        'Price': itemlist_df.iloc[:, 11],
        'Cost': itemlist_df.iloc[:, 14]
    })

    # Stock == 0 이거나 Price == 0 인 행 제거
    df = df[(df['Stock'] != 0) & (df['Price'] != 0)]

    # Cost가 0이면 Price/1.35로 계산
    df['Cost'] = df.apply(
        lambda row: row['Price'] / 1.35 if row['Cost'] == 0 else row['Cost'],
        axis=1
    )

    # G, H, I 열 생성
    df['G'] = df['Stock'] * df['Discount']
    df['H'] = df['Stock'] * df['Price']
    df['I'] = df['Stock'] * df['Cost']

    # 합계 계산
    sum_stock = df['Stock'].sum()
    sum_g = df['G'].sum()
    sum_h = df['H'].sum()
    sum_i = df['I'].sum()

    # total_row 생성
    total_row = {
        'Barcode': '총합계',
        'Description': '',
        'Stock': sum_stock,
        'Discount': '',
        'Price': '',
        'Cost': '',
        'G': sum_g,
        'H': sum_h,
        'I': sum_i
    }

    # df에 새 행 추가 (방법 1: loc 사용)
    df.loc[len(df)] = total_row

    # 최종 결과를 Combined_날짜.xlsx로 저장
    output_filename = f"Stock_{today_str}.xlsx"
    output_path = os.path.join(base_dir, output_filename)
    df.to_excel(output_path, index=False)
    print(f"생성 완료: {output_path}")

if __name__ == "__main__":
    create_combined_excel()
