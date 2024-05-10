import pandas as pd

# 파일 경로 설정
file_paths = {
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\PA.CSV",
}

# 파일 읽기 및 "Total Inventory" 행 찾기
for state, path in file_paths.items():
    # CSV 파일을 데이터프레임으로 로드하여 엑셀 시트에 추가 (헤더를 2로 설정)
    df = pd.read_csv(path, encoding='ISO-8859-1', header=2)

    # "Total Inventory" 행 찾기
    total_inventory_row = df[df.iloc[:, 0] == "Total Inventory"]

    # 행 출력
    if not total_inventory_row.empty:
        print(f"Row containing 'Total Inventory' in {state}:")
        print(total_inventory_row.iloc[0][['Unnamed: 2', 'Unnamed: 5']])
    else:
        print(f"'Total Inventory' not found in {state} data.")
