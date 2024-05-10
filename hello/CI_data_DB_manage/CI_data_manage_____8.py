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

# 4번째 - pcs 열을 기준으로 NaN이 아닌 행만 선택하여 새로운 DataFrame 생성
df_item_remove = df.dropna(subset=[4])

# PO# 주문 섹션의 시작 인덱스 찾기 (df_item_remove 사용)
po_indices = df_item_remove.index[df_item_remove.iloc[:, 0].str.contains("PO#", na=False)].tolist()

# 각 PO# 주문 섹션을 별도의 데이터프레임으로 추출 (df_item_remove 사용)
po_sections_remove = {}
for i, start_index in enumerate(po_indices):
    end_index = po_indices[i + 1] if i + 1 < len(po_indices) else len(df_item_remove)
    section_df = df_item_remove.iloc[start_index:end_index].reset_index(drop=True)
    po_sections_remove[f'po_section_{i}'] = section_df

# 각 PO# 주문 섹션 데이터프레임 출력 (df_item_remove 사용)
for name, section in po_sections_remove.items():
    print(f"{name}:")
    print(section)
    print("\n")
