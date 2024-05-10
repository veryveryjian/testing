import pandas as pd

# 파일 경로 설정
file_paths = {
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY.CSV",
    "NJ": "C:\\Users\\charlton\\Desktop\\COGS\\NJ.CSV",
    "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT.CSV",
    "PA": "C:\\Users\\charlton\\Desktop\\COGS\\PA.CSV"
}

# 엑셀 파일로 데이터프레임 저장
with pd.ExcelWriter('outputxx.xlsx') as writer:
    for state, path in file_paths.items():
        # CSV 파일을 데이터프레임으로 로드하여 엑셀 시트에 추가 (헤더를 0으로 설정)
        df = pd.read_csv(path, encoding='ISO-8859-1', header=0)
        df.to_excel(writer, sheet_name=state)

        # 각 파일의 헤더 출력
        print(f"{state} - Header:")
        print(df.columns.values)  # 헤더 출력
