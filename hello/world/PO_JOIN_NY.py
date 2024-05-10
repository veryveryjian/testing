import pandas as pd

# 엑셀 파일 경로
ny_invoice_file = "C:\\Users\\charlton\\Desktop\\ny_invoice_data.xlsx"
po_nj_file = "C:\\Users\\charlton\\Desktop\\po_nj_data.xlsx"
po_ct_file = "C:\\Users\\charlton\\Desktop\\po_ct_data.xlsx"
po_pa_file = "C:\\Users\\charlton\\Desktop\\po_pa_data.xlsx"
po_mi_file = "C:\\Users\\charlton\\Desktop\\po_mi_data.xlsx"  #만들어여 돼 데이터가 없어서 일단 pass  인보이스만 가져오도록 우선 설정


# 엑셀 파일에서 필요한 데이터 로드 (NJ 데이터)
ny_invoice_data_nj = pd.read_excel(ny_invoice_file, sheet_name='Chunghua NJ', usecols=['Type', 'Date', 'Num', 'Debit',])
po_nj_data = pd.read_excel(po_nj_file, sheet_name='Charlton Cabinetry - NY', usecols=['Type', 'Date', 'Num', 'Credit', 'Memo', 'Memo'])
merged_data_nj = pd.merge(ny_invoice_data_nj, po_nj_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_nj = merged_data_nj[merged_data_nj['_merge'] != 'both']

# 엑셀 파일에서 필요한 데이터 로드 (CT 데이터)
ny_invoice_data_ct = pd.read_excel(ny_invoice_file, sheet_name='Chunghua CT', usecols=['Type', 'Date', 'Num', 'Debit'])
po_ct_data = pd.read_excel(po_ct_file, sheet_name='CHARLTON CABINETRY NY', usecols=['Type', 'Date', 'Num', 'Credit', 'Memo'])
merged_data_ct = pd.merge(ny_invoice_data_ct, po_ct_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_ct = merged_data_ct[merged_data_ct['_merge'] != 'both']

# 엑셀 파일에서 필요한 데이터 로드 (PA 데이터)
ny_invoice_data_pa = pd.read_excel(ny_invoice_file, sheet_name='Charlton PA', usecols=['Type', 'Date', 'Num', 'Debit'])
po_pa_data = pd.read_excel(po_pa_file, sheet_name='CHARLTON NY', usecols=['Type', 'Date', 'Num', 'Credit', 'Memo'])  # 수정된 부분
merged_data_pa = pd.merge(ny_invoice_data_pa, po_pa_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_pa = merged_data_pa[merged_data_pa['_merge'] != 'both']

# 엑셀 파일에서 필요한 데이터 로드 (mi 데이터)
ny_invoice_data_mi = pd.read_excel(ny_invoice_file, sheet_name='Miller Cabinetry', usecols=['Type', 'Date', 'Num', 'Debit'])
#po_mi_data = pd.read_excel(po_pa_file, sheet_name='Charlton Cabinetry Inc', usecols=['Type', 'Date', 'Num', 'Credit'])  # 확인해야 돼
#merged_data_pa = pd.merge(ny_invoice_data_mi, po_mi_data, left_on=['Type', 'Date', 'Num', 'Debit'],
#                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
#unique_values_mi = merged_data_mi[merged_data_mi['_merge'] != 'both']
unique_values_mi = ny_invoice_data_mi

# 결과를 엑셀 파일로 저장
output_file = "C:\\Users\\charlton\\Desktop\\NY.xlsx"
with pd.ExcelWriter(output_file) as writer:
    unique_values_nj.to_excel(writer, sheet_name='ny2nj', index=False)
    unique_values_ct.to_excel(writer, sheet_name='ny2ct', index=False)
    unique_values_pa.to_excel(writer, sheet_name='ny2pa', index=False)
    unique_values_mi.to_excel(writer, sheet_name='ny2mi', index=False)

