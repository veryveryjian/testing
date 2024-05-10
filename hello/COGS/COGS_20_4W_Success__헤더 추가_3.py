import pandas as pd

# 파일 경로 설정
file_paths = {
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY.CSV",
    "NJ": "C:\\Users\\charlton\\Desktop\\COGS\\NJ.CSV",
    "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT.CSV",
    "PA": "C:\\Users\\charlton\\Desktop\\COGS\\PA.CSV"
}

# 두 데이터프레임을 저장할 빈 리스트 생성
dfs1 = []
dfs2 = []

# 결과를 저장할 빈 리스트 생성
inventory_info = []

# 파일 읽기 및 "Total Inventory" 행 찾기
for state, path in file_paths.items():
    # 첫 번째 데이터프레임을 생성하여 엑셀 시트에 추가 (헤더를 2로 설정)
    df1 = pd.read_csv(path, encoding='ISO-8859-1', header=2)

    # "Total Inventory" 행 찾기
    total_inventory_row = df1[df1.iloc[:, 0] == "Total Inventory"]

    # 행 출력
    if not total_inventory_row.empty:
        print(f"Row containing 'Total Inventory' in {state}:")
        print(total_inventory_row.iloc[0][['Unnamed: 2', 'Unnamed: 5']].rename({'Unnamed: 2': 'AMOUNT', 'Unnamed: 5': 'COGS'}))
    else:
        print(f"'Total Inventory' not found in {state} data.")

    # 결과를 포맷팅하여 리스트에 추가
    row_values = total_inventory_row.iloc[0]
    inventory_info.append({'State': state, 'AMOUNT': row_values['Unnamed: 2'], 'COGS': row_values['Unnamed: 5']})

    # 첫 번째 데이터프레임을 리스트에 추가
    dfs1.append(df1)

    # 두 번째 데이터프레임을 생성하여 리스트에 추가
    df2 = pd.read_csv(path, encoding='ISO-8859-1', header=0)
    dfs2.append(df2)

    # 각 파일의 헤더 출력
    print(f"{state} - Header:")
    print(df2.columns.values)  # 헤더 출력

    # 헤더 정보를 결과 리스트에 추가
    header_info = {'State': state, 'Header': df2.columns.values}
    inventory_info.append(header_info)

# 결과를 DataFrame으로 변환
formatted_inventory_df = pd.DataFrame(inventory_info)

# 엑셀 파일로 데이터프레임 저장 (두 시트로)
with pd.ExcelWriter('outputxx.xlsx') as writer:
    formatted_inventory_df.to_excel(writer, sheet_name='Formatted Inventory', index=False)
    for state, df1, df2 in zip(file_paths.keys(), dfs1, dfs2):
        df1.to_excel(writer, sheet_name=f"{state}_Sheet1", index=False)
        df2.to_excel(writer, sheet_name=f"{state}_Sheet2", index=False)
