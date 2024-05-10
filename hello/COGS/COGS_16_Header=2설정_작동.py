import pandas as pd

# 파일 경로 설정
file_paths = {
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY.CSV",
}

# 엑셀 파일로 데이터프레임 저장
with pd.ExcelWriter('outputxx.xlsx') as writer:
    for state, path in file_paths.items():
        # CSV 파일을 데이터프레임으로 로드하여 엑셀 시트에 추가 (헤더를 0으로 설정)
        df = pd.read_csv(path, encoding='ISO-8859-1', header=2)

        # 각 파일의 첫 번째 열과 헤더 출력
        print(f"{state} - First Column:")
        print(df.iloc[:, 0])  # 첫 번째 열 출력

        # 각 파일의 헤더 출력
        print(f"{state} - Header:")
        print(df.columns.values)  # 헤더 출력

        # DataFrame을 Excel 파일의 각 시트에 저장
        df.to_excel(writer, sheet_name=state)

        print("")

        # df1에 저장
        if state == "NY":
            df1 = df.copy()

# df2에 df1의 첫 번째 열만 포함하여 저장
df2 = df1.iloc[:, [0]]

# df2의 첫 번째 열 출력
print("First Column of df2:")
print(df2.iloc[:, 0])

# Get the shape of df2
num_rows, num_columns = df2.shape

print(f"Number of columns in df2: {num_columns}")
print(f"Number of rows in df2: {num_rows}")

# 각각의 컬럼을 분리해서 출력
for column in df2.columns:
    print(f"Column '{column}':")
    print(df2[column])
    print("")

# df2를 Excel 파일로 저장 (기존 파일 덮어쓰기)
df2.to_excel('outputxx.xlsx', sheet_name='df2')
