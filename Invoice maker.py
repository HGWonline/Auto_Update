import pandas as pd
import os

# 현재 스크립트 위치 기준 CSV 파일 경로
csv_filename = "HanGaWeeMarket_Online Invoice.csv"

# CSV 파일 불러오기
df = pd.read_csv(csv_filename)

# 'Order name' 값 기준으로 그룹화 및 개별 파일 생성
unique_orders = df["Order name"].unique()

for order_name in unique_orders:
    df_order = df[df["Order name"] == order_name].copy()

    df_order["Description"] = df_order["Product title"]
    df_order["Quantity"] = df_order["Net items sold"]
    df_order["Unit Price"] = df_order["Product variant compare at price"]
    df_order["Discount Price"] = df_order["Product variant price"]

    # Discount Rate: (Unit Price - Discount Price) / Unit Price * 100
    df_order["Discount Rate"] = ((df_order["Unit Price"] - df_order["Discount Price"]) / df_order["Unit Price"] * 100).round(2)

    # Unit Cost
    df_order["Unit Cost"] = df_order["Cost of goods sold"]

    # Discount Cost = Unit Cost * (1 - Discount Rate / 100)
    df_order["Discount Cost"] = (df_order["Unit Cost"] * (1 - df_order["Discount Rate"] / 100)).round(2)

    # Total Cost = Quantity * Discount Cost
    df_order["Total Cost"] = (df_order["Quantity"] * df_order["Discount Cost"]).round(2)

    # 인보이스용 데이터프레임
    invoice_df = df_order[
        ["Description", "Quantity", "Unit Price", "Discount Price", "Discount Rate",
         "Total sales", "Unit Cost", "Discount Cost", "Total Cost"]
    ]

    # 엑셀 파일로 저장 (예: invoice_order_1020.xlsx)
    order_id = order_name.replace("#", "")  # 파일명에서 '#' 제거
    output_filename = f"invoice_order_{order_id}.xlsx"
    invoice_df.to_excel(output_filename, index=False)
    print(f"{output_filename} 생성 완료.")
