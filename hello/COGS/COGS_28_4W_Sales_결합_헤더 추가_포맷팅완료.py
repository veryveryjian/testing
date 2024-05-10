import pandas as pd

# 파일 경로 설정
file_paths = {
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY_SALES.CSV",
    "NJ": "C:\\Users\\charlton\\Desktop\\COGS\\NJ_SALES.CSV",
    "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT_SALES.CSV",
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
            results.append({'State': state, 'Value': inventory_value, 'First Column': ''})
        else:
            results.append({'State': state, 'Value': None, 'First Column': ''})
    else:
        if state == "NJ":
            last_row_value = df.iloc[-1]['Unnamed: 10']
        elif state == "NY":
            last_row_value = df.iloc[-1]['Unnamed: 11']
        elif state == "PA":
            last_row_value = df.iloc[-1]['Unnamed: 11']

        last_row_first_col_value = df.iloc[-1, 0]
        results.append({'State': state, 'Value': last_row_value, 'First Column': last_row_first_col_value})

# 결과 데이터프레임 생성
results_df = pd.DataFrame(results)

# 순서대로 정렬 (NY, NJ, CT, PA 순서)
ordered_states = ['NY', 'NJ', 'CT', 'PA']
results_df = results_df.set_index('State').reindex(ordered_states).reset_index()

# 데이터프레임의 컬럼 순서를 정리
results_df = results_df[['State', 'Value', 'First Column']]

# 결과 출력
print(results_df)

# 엑셀 파일로 저장
results_df.to_excel("output_sales_data_formatted.xlsx", index=False)
