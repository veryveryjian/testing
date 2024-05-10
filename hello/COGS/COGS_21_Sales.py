import pandas as pd

# pandas 설정 변경: DataFrame의 모든 행과 열을 출력하도록 설정
pd.set_option('display.max_rows', None)  # 모든 행 출력
pd.set_option('display.max_columns', None)  # 모든 열 출력
pd.set_option('display.width', None)  # 셀 너비 제한 없애기
pd.set_option('display.max_colwidth', None)  # 컬럼 내용을 생략 없이 전체 출력

# 파일 경로 설정
file_paths = {
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY_SALES.CSV",
    "NJ": "C:\\Users\\charlton\\Desktop\\COGS\\NJ_SALES.CSV",
    "PA": "C:\\Users\\charlton\\Desktop\\COGS\\PA_SALES.CSV",
    # "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT_SALES.CSV",
}

# 파일 읽기 및 데이터 출력
for state, path in file_paths.items():
    # CSV 파일을 데이터프레임으로 로드
    df = pd.read_csv(path, encoding='ISO-8859-1', header=1)

    # 데이터 출력
    # print(f"Data for {state}:")
    # print(df)

    # 마지막 한 줄 출력
    print(f"Last row for {state}:")
    print(df.tail(1))
