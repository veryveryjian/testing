import pandas as pd

# CSV 파일 경로 설정
csv_path = "C:/Users/charlton/Desktop/COGS_/CT.csv"

# CSV 파일 읽기, 헤더가 없다고 지정
df = pd.read_csv(csv_path, header=None)

# 'Total Inventory' 문자열이 포함된 행 찾기
inventory_rows = df[df[0] == 'Total Inventory'].index

if not inventory_rows.empty:
    inventory_row_index = inventory_rows[0]
    # 2번째 열과 5번째 열의 값을 선택
    second_column_value = df.iloc[inventory_row_index, 2]
    fifth_column_value = df.iloc[inventory_row_index, 5]

    # 결과 출력
    print("AMOUNT:", second_column_value)
    print("COGS:", fifth_column_value)
else:
    print("'Total Inventory'에 해당하는 데이터가 없습니다.")
