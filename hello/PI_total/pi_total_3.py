import os
import pandas as pd

# 경로 설정
directory = r'C:\Users\charlton\Desktop\rere'
output_directory = r'C:\Users\charlton\Desktop\rere\output'

# 출력 디렉토리 생성 (없으면 생성)
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# 해당 디렉토리의 모든 파일을 순회
for filename in os.listdir(directory):
    if filename.endswith(".xls") or filename.endswith(".xlsx"):
        # 파일의 전체 경로를 구성
        filepath = os.path.join(directory, filename)

        try:
            # 엑셀 파일 읽기
            df = pd.read_excel(filepath)

            # 'TOTAL' 문자열을 포함하는 행 찾기
            total_rows = df[df.iloc[:, 0].astype(str).str.contains("TOTAL")]

            # 데이터가 있는 경우만 파일로 저장
            if not total_rows.empty:
                # 새 파일명 설정
                new_filename = filename.replace(".xls", "_filtered.xlsx").replace(".xlsx", "_filtered.xlsx")
                output_filepath = os.path.join(output_directory, new_filename)

                # 엑셀 파일로 저장
                total_rows.to_excel(output_filepath, index=False)
                print(f"'TOTAL' 행을 포함하는 데이터가 {output_filepath}에 저장되었습니다.")

        except Exception as e:
            print(f"파일을 처리하는 중 오류 발생: {filename}")
            print(str(e))

