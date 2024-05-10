import pandas as pd

# 엑셀 파일 경로
nj_invoice_file = "C:\\Users\\charlton\\Desktop\\nj_invoice_data.xlsx"
po_ny_file = "C:\\Users\\charlton\\Desktop\\po_ny_data.xlsx"
po_ct_file = "C:\\Users\\charlton\\Desktop\\po_ct_data.xlsx"
po_pa_file = "C:\\Users\\charlton\\Desktop\\po_pa_data.xlsx"  # 수정된 부분
po_mi_file = "C:\\Users\\charlton\\Desktop\\po_mi_data.xlsx"  #만들어여 돼 데이터가 없어서 일단 pass  인보이스만 가져오도록 우선 설정


# 엑셀 파일에서 필요한 데이터 로드 (NJ 데이터)
nj_invoice_data_nj = pd.read_excel(nj_invoice_file, sheet_name='Charlton Cabinetry', usecols=['Type', 'Date', 'Num', 'Debit', ])
po_ny_data = pd.read_excel(po_ny_file, sheet_name='CHUNG HUA CABINET NJ INC', usecols=['Type', 'Date', 'Num', 'Credit', 'Memo'])
merged_data_ny = pd.merge(nj_invoice_data_nj, po_ny_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_nj = merged_data_ny[merged_data_ny['_merge'] != 'both']

# 엑셀 파일에서 필요한 데이터 로드 (CT 데이터)
ny_invoice_data_ct = pd.read_excel(nj_invoice_file, sheet_name='Chunghua Cabinet CT Inc.', usecols=['Type', 'Date', 'Num', 'Debit', ])
po_ct_data = pd.read_excel(po_ct_file, sheet_name='Chung Hua Cabinet NJ Inc', usecols=['Type', 'Date', 'Num', 'Credit', 'Memo'])
merged_data_ct = pd.merge(ny_invoice_data_ct, po_ct_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_ct = merged_data_ct[merged_data_ct['_merge'] != 'both']

# 엑셀 파일에서 필요한 데이터 로드 (PA 데이터)
ny_invoice_data_pa = pd.read_excel(nj_invoice_file, sheet_name='Charlton Cabinetry -- PA', usecols=['Type', 'Date', 'Num', 'Debit', ])
po_pa_data = pd.read_excel(po_pa_file, sheet_name='CHARLTON NJ', usecols=['Type', 'Date', 'Num', 'Credit', 'Memo'])  # 수정된 부분
merged_data_pa = pd.merge(ny_invoice_data_pa, po_pa_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_pa = merged_data_pa[merged_data_pa['_merge'] != 'both']


#mi inv
ny_invoice_data_mi = pd.read_excel(nj_invoice_file, sheet_name='Miller Cabinetry Inc.', usecols=['Type', 'Date', 'Num', 'Debit'])
unique_values_mi = ny_invoice_data_mi



# 결과를 엑셀 파일로 저장
output_file = "C:\\Users\\charlton\\Desktop\\NJ.xlsx"
with pd.ExcelWriter(output_file) as writer:
    unique_values_nj.to_excel(writer, sheet_name='nj2ny', index=False)
    unique_values_ct.to_excel(writer, sheet_name='nj2ct', index=False)
    unique_values_pa.to_excel(writer, sheet_name='nj2pa', index=False)
    unique_values_mi.to_excel(writer, sheet_name='nj2mi', index=False)

