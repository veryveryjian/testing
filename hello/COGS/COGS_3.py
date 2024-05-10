import pandas as pd

# 파일 경로 설정
file_path = "C:/Users/charlton/Desktop/COGS/NY.CSV"

# pandas 설정 변경: DataFrame의 모든 행과 열을 출력하도록 설정
pd.set_option('display.max_rows', None)  # 모든 행 출력
pd.set_option('display.max_columns', None)  # 모든 열 출력
pd.set_option('display.width', None)  # 셀 너비 제한 없애기
pd.set_option('display.max_colwidth', None)  # 컬럼 내용을 생략 없이 전체 출력

# 파일의 첫 5줄을 읽어 구분자 확인
with open(file_path, 'r') as file:
    for _ in range(5):
        print(file.readline())

# 데이터를 올바른 구분자로 읽어보기 (예: 탭 '\t')
try:
    # 헤더가 첫 번째 줄일 때 (header=0)
    print("헤더가 첫 번째 줄일 때 (header=0):")
    df_header_0 = pd.read_csv(file_path, header=0, sep='\t')  # 구분자를 탭으로 설정
    print(df_header_0.head())
    print("\n")

    # 헤더가 없을 때 (header=None)
    print("헤더가 없을 때 (header=None):")
    df_header_none = pd.read_csv(file_path, header=None, sep='\t')  # 구분자를 탭으로 설정
    print(df_header_none.head())
    print("\n")

    # 헤더가 두 번째 줄일 때 (header=1)
    print("헤더가 두 번째 줄일 때 (header=1):")
    df_header_1 = pd.read_csv(file_path, header=1, sep='\t')  # 구분자를 탭으로 설정
    print(df_header_1.head())
    print("\n")

    # 헤더가 세 번째 줄일 때 (header=2)
    print("헤더가 세 번째 줄일 때 (header=2):")
    df_header_2 = pd.read_csv(file_path, header=2, sep='\t')  # 구분자를 탭으로 설정
    print(df_header_2.head())
    print("\n")

except pd.errors.ParserError as e:
    print(f"Parser Error: {e}")
except Exception as e:
    print(f"An error occurred: {e}")
