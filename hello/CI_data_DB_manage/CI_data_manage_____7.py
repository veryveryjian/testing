import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# Excel 설정
excel_file = r'C:\Users\charlton\Desktop\test1.xlsx'
df = pd.read_excel(excel_file, sheet_name='INVOICE', skiprows=14, usecols='A:G', header=None)

# 첫 번째 행을 헤더로 설정 및 데이터 전처리
df = df.drop(0).reset_index(drop=True)  # 첫 번째 행 제거 후 인덱스 재설정

# PO# 주문 섹션의 시작 인덱스 찾기
po_indices = df.index[df.iloc[:, 0].str.contains("PO#", na=False)].tolist()

# 각 PO# 주문 섹션을 별도의 데이터프레임으로 추출
po_sections = {}
for i, start_index in enumerate(po_indices):
    end_index = po_indices[i + 1] if i + 1 < len(po_indices) else len(df)
    section_df = df.iloc[start_index:end_index].reset_index(drop=True)
    po_sections[f'po_section_{i}'] = section_df
# [***************pcs*****************]
# 1번째 열에 데이터가 있는 행만 필터링하여 새로운 DataFrame 생성
df_item_remove = df.dropna(subset=[4])  # 4번째 - pcs 열을 기준으로 NaN이 아닌 행만 선택C

# 결과 확인
print(df_item_remove)
#이걸로 활용하는것에서 문제가 생김 특히 페이짼부훠에서

