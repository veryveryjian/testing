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

# 결과 출력
print(df2)
