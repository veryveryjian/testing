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

    # 두 번째 데이터프레임을 생성하여 리스트에 추가
    df2 = pd.read_csv(path, encoding='ISO-8859-1', header=0)

    # 결과 딕셔너리 생성
    inventory_dict = {'State': state}

    # "Total Inventory" 행이 있으면 그 값을 딕셔너리에 추가
    if not total_inventory_row.empty:
        row_values = total_inventory_row.iloc[0]
        inventory_dict['AMOUNT'] = row_values['Unnamed: 2']
        inventory_dict['COGS'] = row_values['Unnamed: 5']
    else:
        inventory_dict['AMOUNT'] = None
        inventory_dict['COGS'] = None

    # 헤더 정보를 딕셔너리에 추가
    inventory_dict['Header'] = df2.columns.values

    # 최종 정보를 inventory_info 리스트에 추가
    inventory_info.append(inventory_dict)

    # 데이터프레임들을 리스트에 추가
    dfs1.append(df1)
    dfs2.append(df2)

# 결과를 DataFrame으로 변환
formatted_inventory_df = pd.DataFrame(inventory_info)

# 엑셀 파일로 데이터프레임 저장 (두 시트로)
with pd.ExcelWriter('outputxx.xlsx') as writer:
    formatted_inventory_df.to_excel(writer, sheet_name='Formatted Inventory', index=False)
    for state, df1, df2 in zip(file_paths.keys(), dfs1, dfs2):
        df1.to_excel(writer, sheet_name=f"{state}_Sheet1", index=False)
        df2.to_excel(writer, sheet_name=f"{state}_Sheet2", index=False)

print(formatted_inventory_df)