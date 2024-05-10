import pandas as pd

# 파일 경로 설정
file_path = "C:/Users/charlton/Desktop/COGS/NY.CSV"

# pandas 설정 변경: DataFrame의 모든 행과 열을 출력하도록 설정
pd.set_option('display.max_rows', None)  # 모든 행 출력
pd.set_option('display.max_columns', None)  # 모든 열 출력
pd.set_option('display.width', None)  # 셀 너비 제한 없애기
pd.set_option('display.max_colwidth', None)  # 컬럼 내용을 생략 없이 전체 출력

# 올바른 구분자로 데이터 로드 (예시에서는 탭 '\t'로 가정)
try:
    # 헤더가 첫 번째 줄일 때 데이터 로드
    df = pd.read_csv(file_path, header=0, sep='\t')

    # 'TOTAL' 행과 'Total Inventory' 행의 인덱스 찾기
    total_row_index = df.index[df.iloc[:, 0] == "TOTAL"].tolist()[0]
    total_inventory_row_index = df.index[df.iloc[:, 0] == "Total Inventory"].tolist()[0]

    # 'May 8, 24', 'Amount', 'COGS' 열의 인덱스 찾기
    columns = df.columns
    may_8_24_col_index = columns.get_loc("May 8, 24")
    amount_col_index = columns.get_loc("Amount")
    cogs_col_index = columns.get_loc("COGS")

    # 'May 8, 24' 열의 'TOTAL' 행 값 추출
    may_8_24_total = df.iloc[total_row_index, may_8_24_col_index]

    # 'Amount'와 'COGS' 열의 'Total Inventory' 행 값 추출
    amount_total_inventory = df.iloc[total_inventory_row_index, amount_col_index]
    cogs_total_inventory = df.iloc[total_inventory_row_index, cogs_col_index]

    # 결과 출력
    print("\nSelected Values:")
    print("[Row, Column] for May 8, 24 (TOTAL 행):", [total_row_index, may_8_24_col_index], "=", may_8_24_total)
    print("[Row, Column] for Amount (Total Inventory 행):", [total_inventory_row_index, amount_col_index], "=", amount_total_inventory)
    print("[Row, Column] for COGS (Total Inventory 행):", [total_inventory_row_index, cogs_col_index], "=", cogs_total_inventory)

except pd.errors.ParserError as e:
    print(f"Parser Error: {e}")
except Exception as e:
    print(f"An error occurred: {e}")
