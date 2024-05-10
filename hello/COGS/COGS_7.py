import pandas as pd

# pandas 설정 변경: DataFrame의 모든 행과 열을 출력하도록 설정
pd.set_option('display.max_rows', None)  # 모든 행 출력
pd.set_option('display.max_columns', None)  # 모든 열 출력
pd.set_option('display.width', None)  # 셀 너비 제한 없애기
pd.set_option('display.max_colwidth', None)  # 컬럼 내용을 생략 없이 전체 출력

# CSV 파일 경로 설정
csv_path = "C:/Users/charlton/Desktop/COGS/NY.CSV"

# 파일의 첫 5줄을 읽어 출력하기 (디버깅 목적)
with open(csv_path, 'r') as file:
    for _ in range(5):
        print(file.readline())

# 구분자를 확인한 후 적절한 구분자로 CSV 파일 읽기
try:
    df = pd.read_csv(csv_path, header=None, sep=',')
except pd.errors.ParserError as e:
    print("ParserError:", e)
    print("구분자 또는 파일 형식을 확인하세요.")
    # 예외 발생 시 파일을 다시 읽어보는 로직 (다른 구분자를 시도)
    try:
        df = pd.read_csv(csv_path, header=None, sep='\t')
    except Exception as e:
        print("다른 오류:", e)
        print("파일 읽기에 실패했습니다.")
        exit()

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
