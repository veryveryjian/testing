import os
import pandas as pd

# 네트워크 드라이브 경로 설정
path = r"\\NAS132B9F\1.Purchase order _订单&统计\1-8订单\1-8-23越南三洋订单\PO"

# 해당 경로에 있는 파일들을 리스트로 가져오기
files = os.listdir(path)

# 엑셀 파일만 필터링하여 파일 이름 저장
excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xls')]

# 파일 이름을 데이터프레임에 저장
PO_df = pd.DataFrame(excel_files, columns=['File Name'])

# 파일 이름에서 첫 번째 하이픈('-') 앞의 내용만 추출하여 새로운 데이터프레임에 저장
df2 = PO_df['File Name'].apply(lambda x: x.split('-')[0]).to_frame(name='Prefix')

# 'x' 및 '#' 문자열 제거 후 새로운 데이터프레임 생성
df3 = df2['Prefix'].replace(['x', '#'], '', regex=True).to_frame(name='Prefix_no_x_hash')

# 숫자로 변환을 시도하고 실패할 경우 NaN 반환
df4 = df3['Prefix_no_x_hash'].apply(pd.to_numeric, errors='coerce').to_frame(name='Integer_Values')

# 오름차순으로 정렬
df4_sorted = df4.sort_values(by='Integer_Values')

# 바탕화면 경로 설정 (윈도우 사용자의 경우)
desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
output_file_path = os.path.join(desktop_path, 'Extracted_Prefixes_po.xlsx')

# 데이터프레임을 엑셀 파일로 저장 - 각 데이터프레임을 별도의 시트에 저장
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    df4_sorted.to_excel(writer, sheet_name='Sorted_Integer_Values', index=False)

print(f"파일이 저장되었습니다: {output_file_path}")
