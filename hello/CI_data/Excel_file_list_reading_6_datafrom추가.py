# datafrom 컬럼을 추가

import os
import pandas as pd

# 디렉터리 설정: 바탕화면에 CI_Create 폴더 위치
directory = r'C:\Users\charlton\Desktop\CI_data_r3'
output_directory = r'C:\Users\charlton\Desktop\CI_Create'

# CI_Create 폴더가 없다면 생성
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# 디렉터리 내의 모든 Excel 파일 이름을 리스트로 저장
excel_files = [filename for filename in os.listdir(directory) if filename.endswith('.xlsx')]

# 각 Excel 파일의 "INVOICE" 시트를 읽고 새 파일로 저장
for file in excel_files:
    file_path = os.path.join(directory, file)
    try:
        # "INVOICE" 시트만 읽기, skiprows=14 적용
        invoice_df = pd.read_excel(file_path, sheet_name='INVOICE', skiprows=0)

        # 열 1부터 6까지 모든 데이터가 NaN인 행 제거
        invoice_df.dropna(subset=invoice_df.columns[1:7], how='all', inplace=True)

        # 'DataFrom' 컬럼 추가, 원본 파일 이름으로 모든 행 채우기
        invoice_df['DataFrom'] = file

        # 새 파일 경로 설정 (원본 파일 이름 사용)
        new_file_path = os.path.join(output_directory, f"{os.path.splitext(file)[0]}_INVOICE_f2.xlsx")

        # 데이터를 새 Excel 파일로 저장
        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            invoice_df.to_excel(writer, index=False, sheet_name='INVOICE')

        print(f"Processed and saved: {new_file_path}")

    except Exception as e:
        print(f"Failed to process {file}: {e}")
