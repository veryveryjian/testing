import pandas as pd

# 파일 경로 설정
file_paths = {
    "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT_SALES.CSV",
}

# 파일 읽기 및 데이터 출력
for state, path in file_paths.items():
    # CSV 파일을 데이터프레임으로 로드
    df = pd.read_csv(path, encoding='ISO-8859-1', header=1)

    # "Total Inventory" 행 찾기
    total_inventory_row = df[df.iloc[:, 0] == "Total Inventory"]

    # "Total Inventory" 행이 존재할 때
    if not total_inventory_row.empty:
        # 3번째 열의 값을 출력
        inventory_value = total_inventory_row.iloc[0, 2]
        print(f"Value of the 3rd column in 'Total Inventory' row for {state}: {inventory_value}")
    else:
        print(f"'Total Inventory' row not found for {state}.")
