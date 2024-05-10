import os
import pandas as pd

# 경로 설정
directory = r'C:\Users\charlton\Desktop\rere'

# 디렉토리 존재 여부 확인
if not os.path.exists(directory):
    print(f"경로가 존재하지 않습니다: {directory}")
else:
    print(f"처리할 디렉토리: {directory}")

# 해당 디렉토리의 모든 파일을 순회
found_excel_files = False

for filename in os.listdir(directory):
    if filename.endswith(".xls") or filename.endswith(".xlsx"):
        found_excel_files = True
        # 파일의 전체 경로를 구성
        filepath = os.path.join(directory, filename)

        try:
            # 엑셀 파일 읽기
            df = pd.read_excel(filepath)

            # 데이터 출력
            print(f"File: {filename}")
            # 첫 번째 열의 데이터 출력
            first_column = df.iloc[:, 0]
            print("첫 번째 열의 데이터:")
            print(first_column)
            print("\n")

            # 'TOTAL' 문자열을 포함하는 행 찾기
            print("'TOTAL' 문자열이 포함된 행:")
            total_rows = df[df.iloc[:, 0].astype(str).str.contains("TOTAL")]
            print(total_rows)
            print("\n" * 2)

        except Exception as e:
            print(f"파일을 읽는 중 오류 발생: {filename}")
            print(str(e))
            print("\n")

if not found_excel_files:
    print("엑셀 파일을 찾을 수 없습니다. '.xls' 또는 '.xlsx' 확장자를 가진 파일이 있는지 확인하세요.")
