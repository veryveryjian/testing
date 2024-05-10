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

    # 첫 번째 데이터프레임을 리스트에 추가
    dfs1.append(df1)

    # 두 번째 데이터프레임을 생성하여 리스트에 추가
    df2 = pd.read_csv(path, encoding='ISO-8859-1', header=0)
    dfs2.append(df2)

    # 각 파일의 헤더 출력
    print(f"{state} - Header:")
    print(df2.columns.values)  # 헤더 출력

# 엑셀 파일로 데이터프레임 저장 (두 시트로)
with pd.ExcelWriter('outputxx.xlsx') as writer:
    for state, df1, df2 in zip(file_paths.keys(), dfs1, dfs2):
        df1.to_excel(writer, sheet_name=f"{state}_Sheet1", index=False)
        df2.to_excel(writer, sheet_name=f"{state}_Sheet2", index=False)
