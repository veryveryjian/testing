import pandas as pd

# 파일 경로 설정
file_paths = {
    "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT_SALES.CSV",
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY_SALES.CSV",
    "NJ": "C:\\Users\\charlton\\Desktop\\COGS\\NJ_SALES.CSV",
    "PA": "C:\\Users\\charlton\\Desktop\\COGS\\PA_SALES.CSV",
}

# 파일 읽기 및 데이터 출력
for state, path in file_paths.items():
    # CSV 파일을 데이터프레임으로 로드
    df = pd.read_csv(path, encoding='ISO-8859-1', header=1)

    # CT 파일인 경우 "Total Inventory" 행의 3번째 열의 값을 출력
    if state == "CT":
        total_inventory_row = df[df.iloc[:, 0] == "Total Inventory"]
        if not total_inventory_row.empty:
            inventory_value = total_inventory_row.iloc[0, 2]
            print(f"Value of last row for {state}: {inventory_value}")
        else:
            print(f"'Total Inventory' row not found for {state}.")

    # NJ, NY, PA 파일인 경우 마지막 행의 값을 출력
    else:
        if state == "NJ":
            last_row_value = df.iloc[-1]['Unnamed: 10']
        else:
            last_row_value = df.iloc[-1]['Unnamed: 11']
        print(f"Value of last row for {state}: {last_row_value}")
