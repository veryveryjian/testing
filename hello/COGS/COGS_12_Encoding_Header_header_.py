import pandas as pd

# 파일 경로 설정
file_paths = {
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY.CSV",
    "NJ": "C:\\Users\\charlton\\Desktop\\COGS\\NJ.CSV",
    "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT.CSV",
    "PA": "C:\\Users\\charlton\\Desktop\\COGS\\PA.CSV"
}

# 각 파일에서 특정 위치의 위에 있는 값을 추출하는 함수
def extract_value_above_position(file_path, row_index, col_index):
    # 파일을 데이터프레임으로 로드 (헤더를 0으로 설정)
    df = pd.read_csv(file_path, encoding='ISO-8859-1', header=0)

    # 지정된 행과 열의 위치에서 위에 있는 값을 추출 (row_index - 1)
    value = df.iloc[row_index - 1, col_index]

    return value

# 엑셀 파일로 데이터프레임 저장
with pd.ExcelWriter('outputxx.xlsx') as writer:
    for state, path in file_paths.items():
        # CSV 파일을 데이터프레임으로 로드하여 엑셀 시트에 추가 (헤더를 0으로 설정)
        df = pd.read_csv(path, encoding='ISO-8859-1', header=0)
        df.to_excel(writer, sheet_name=state)

        # 각 파일의 첫 번째 열과 헤더 출력
        print(f"{state} - First Column:")
        # print(df.iloc[:, 0])  # 첫 번째 열 출력
        print(f"{state} - Header:")
        print(df.columns.values)  # 헤더 출력
        print("")

        # 각 파일에서 특정 위치의 위에 있는 값을 추출하고 출력
        value = extract_value_above_position(path, 2, 1)
        print(f"{state} - Value above position : {value}")
