import pandas as pd

# 출력 제한 해제
pd.set_option('display.max_rows', None)  # 행 제한 해제
pd.set_option('display.max_columns', None)  # 열 제한 해제
pd.set_option('display.width', None)  # 너비 제한 해제
pd.set_option('display.max_colwidth', None)  # 열 내용 너비 제한 해제


# Excel 파일 경로
excel_file = r'C:\Users\charlton\Desktop\test1.xlsx'

# 'INVOICE' 시트에서 B16:G 범위의 데이터를 읽어오기
df = pd.read_excel(excel_file, sheet_name='INVOICE', skiprows=15, usecols='B:G')

# 결과를 새 Excel 파일로 저장
output_file = r'C:\Users\charlton\Desktop\filtered_data.xlsx'
df.to_excel(output_file, index=False)

print(df)

print("데이터가 성공적으로 저장되었습니다:", output_file)



