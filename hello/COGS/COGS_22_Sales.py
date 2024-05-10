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

    # 데이터 출력
    # print(f"Data for {state}:")
    # print(df)

    # 마지막 줄 출력
    print(f"Last row for {state}:")
    print(df.tail(1))

    # 마지막 열의 값을 가져오기
    last_column_value = df.iloc[-1, -1]
    print(f"Last column value for {state}: {last_column_value}")
