import pandas as pd

# 파일 경로 설정
file_paths = {
    "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT_SALES.CSV",
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY_SALES.CSV",
    "NJ": "C:\\Users\\charlton\\Desktop\\COGS\\NJ_SALES.CSV",
    "PA": "C:\\Users\\charlton\\Desktop\\COGS\\PA_SALES.CSV",
}

# 결과를 저장할 리스트
results = []

# 파일 읽기 및 데이터 출력
for state, path in file_paths.items():
    # CSV 파일을 데이터프레임으로 로드
    df = pd.read_csv(path, encoding='ISO-8859-1', header=1)

    # 데이터 추출 및 저장
    if state == "CT":
        total_inventory_row = df[df.iloc[:, 0] == "Total Inventory"]
        if not total_inventory_row.empty:
            inventory_value = total_inventory_row.iloc[0, 2]
            results.append({'State': state, 'Type': 'Total Inventory', 'Value': inventory_value})
        else:
            results.append({'State': state, 'Type': 'Total Inventory', 'Value': None})
    else:
        if state == "NJ":
            last_row_value = df.iloc[-1]['Unnamed: 10']
        elif state == "NY":
            last_row_value = df.iloc[-1]['Unnamed: 11']
        elif state == "PA":
            last_row_value = df.iloc[-1]['Unnamed: 11']

        results.append({'State': state, 'Type': 'Last Row Specific Column', 'Value': last_row_value})

        last_row_first_col_value = df.iloc[-1, 0]
        results.append({'State': state, 'Type': 'Last Row First Column', 'Value': last_row_first_col_value})

# 결과 데이터프레임 생성
results_df = pd.DataFrame(results)

# 결과 출력
print(results_df)

# 엑셀 파일로 저장
results_df.to_excel("output_sales_data.xlsx", index=False)
