import os
import pandas as pd

# 네트워크 드라이브 경로 설정
path = r"\\NAS132B9F\1.Purchase order _订单&统计\1-8订单\1-8-22 越南三洋—出货单"

# 해당 경로에 있는 파일들을 리스트로 가져오기
files = os.listdir(path)

# 엑셀 파일만 필터링하여 파일 이름 저장
excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xls')]

# 파일 이름을 데이터프레임에 저장
PO_df = pd.DataFrame(excel_files, columns=['File Name'])

# 파일 이름에서 "PO#" 뒤의 내용 추출하여 새로운 데이터프레임에 저장
def extract_prefix(file_name):
    if 'PO#' in file_name:
        return file_name.split('PO#')[1].split('.')[0].strip()
    else:
        return file_name

df2 = PO_df['File Name'].apply(extract_prefix).to_frame(name='Prefix')

# 'INVOICE' 문자열 이후 내용 제거
df3 = df2['Prefix'].apply(lambda x: x.split(' INVOICE')[0] if ' INVOICE' in x else x).to_frame(name='Cleaned_Prefix')

# 바탕화면 경로 설정 (윈도우 사용자의 경우)
desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
output_file_path = os.path.join(desktop_path, 'Extracted_Prefixes_ci.xlsx')

# 데이터프레임을 엑셀 파일로 저장 - 각 데이터프레임을 별도의 시트에 저장
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    PO_df.to_excel(writer, sheet_name='Original_File_Names', index=False)
    df2.to_excel(writer, sheet_name='Prefix_Extracted', index=False)
    df3.to_excel(writer, sheet_name='Cleaned_Prefix', index=False)

print(f"파일이 저장되었습니다: {output_file_path}")
