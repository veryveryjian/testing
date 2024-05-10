import pandas as pd

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

    # # 데이터 출력
    # print(f"Data for {state}:")
    # print(df)
    #
    # # 마지막 줄 출력
    # print(f"Last row for {state}:")
    # print(df.tail(1))

    # 마지막 행 및 첫번째 열 데이터 값 출력
    last_row_first_col_value = df.iloc[-1, 0]

    # NJ의 경우는 Unnamed: 10 열의 데이터 값 출력
    if state == "NJ":
        last_row_value = df.iloc[-1]['Unnamed: 11']
    else:
        # 기타 경우는 Unnamed: 11 열의 데이터 값과 결합하여 출력
        last_row_value = df.iloc[-1]['Unnamed: 11']

    print(f"Value of last row for {state} (first column): {last_row_first_col_value}")
    print(f"Value of last row for {state} (eleventh column): {last_row_value}")
