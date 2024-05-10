import pandas as pd

# 파일 경로 설정
file_path = "C:/Users/charlton/Desktop/COGS/NY.CSV"

# pandas 설정 변경: DataFrame의 모든 행과 열을 출력하도록 설정
pd.set_option('display.max_rows', None)  # 모든 행 출력
pd.set_option('display.max_columns', None)  # 모든 열 출력
pd.set_option('display.width', None)  # 셀 너비 제한 없애기
pd.set_option('display.max_colwidth', None)  # 컬럼 내용을 생략 없이 전체 출력

# 파일의 첫 5줄을 읽어 구분자 확인
print("파일의 첫 5줄:")
with open(file_path, 'r') as file:
    for _ in range(5):
        print(file.readline())

# 올바른 구분자로 데이터 로드 (예시에서는 탭 '\t'로 가정)
# 헤더가 첫 번째 줄인 경우 (일반적인 상황)
print("\n헤더를 제외한 데이터 출력 (header=0):")
try:
    # 헤더가 첫 번째 줄일 때 데이터 로드
    df = pd.read_csv(file_path, header=0, sep='\t')

    # 헤더를 제외하고 데이터 출력
    if not df.empty:
        print(df.iloc[1:])  # 첫 번째 행(헤더 다음 행)부터 출력
    else:
        print("데이터가 없습니다.")

except pd.errors.ParserError as e:
    print(f"Parser Error: {e}")
except Exception as e:
    print(f"An error occurred: {e}")
