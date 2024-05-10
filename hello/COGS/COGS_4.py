import pandas as pd

# 파일 경로 설정
file_path = "C:/Users/charlton/Desktop/COGS/NY.CSV"
output_excel_path = "C:/Users/charlton/Desktop/COGS/Output.xlsx"

# pandas 설정 변경: DataFrame의 모든 행과 열을 출력하도록 설정
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# 파일을 열어 첫 번째 줄 확인
with open(file_path, 'r') as file:
    first_line = file.readline()
print("파일의 첫 줄:", first_line)

# 추정되는 구분자를 기반으로 데이터 로드
# 여기서는 탭(\t), 쉼표(,), 세미콜론(;) 등을 확인
try:
    if '\t' in first_line:
        delimiter = '\t'
    elif ';' in first_line:
        delimiter = ';'
    else:
        delimiter = ','  # 가장 일반적인 구분자

    # 데이터 로드
    df = pd.read_csv(file_path, delimiter=delimiter)

    # 데이터 프레임 출력
    print("\n데이터 프레임:")
    print(df.head())

    # Excel 파일로 저장
    df.to_excel(output_excel_path, index=False)
    print(f"\n데이터가 성공적으로 Excel 파일에 저장되었습니다: {output_excel_path}")

except pd.errors.ParserError as e:
    print(f"Parser Error: {e}")
except Exception as e:
    print(f"An error occurred: {e}")
