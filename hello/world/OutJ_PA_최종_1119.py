import pandas as pd

# 엑셀 파일 경로
pa_invoice_file = "/Users/mymac/Desktop/pa_invoice_data.xlsx"
po_ny_file = "/Users/mymac/Desktop/po_ny_data.xlsx"
po_nj_file = "/Users/mymac/Desktop/po_nj_data.xlsx"
po_ct_file = "/Users/mymac/Desktop/po_ct_data.xlsx"

# 'Chunghua PA' 시트에서 데이터 로드
po_ny_data1 = pd.read_excel(po_ny_file, sheet_name='Chunghua PA', usecols=['Type', 'Date', 'Num', 'Credit'])

# 'Charlton PA INC' 시트에서 데이터 로드
po_ny_data2 = pd.read_excel(po_ny_file, sheet_name='Charlton PA INC', usecols=['Type', 'Date', 'Num', 'Credit'])

# 두 시트의 데이터를 결합
po_ny_data = pd.concat([po_ny_data1, po_ny_data2], ignore_index=True)

# 엑셀 파일에서 필요한 데이터 로드 (NJ 데이터)
pa_invoice_data_nj = pd.read_excel(pa_invoice_file, sheet_name='CHARLTON CABINETRY NY', usecols=['Type', 'Date', 'Num', 'Debit'])
merged_data_nj = pd.merge(pa_invoice_data_nj, po_ny_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_nj = merged_data_nj[merged_data_nj['_merge'] != 'both']

# 엑셀 파일에서 필요한 데이터 로드 (CT 데이터)
pa_invoice_data_ct = pd.read_excel(pa_invoice_file, sheet_name='CHARLTON CABINETRY NJ', usecols=['Type', 'Date', 'Num', 'Debit'])
po_nj_data = pd.read_excel(po_nj_file, sheet_name='Charlton Cabinetry - PA', usecols=['Type', 'Date', 'Num', 'Credit'])
merged_data_ct = pd.merge(pa_invoice_data_ct, po_nj_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_ct = merged_data_ct[merged_data_ct['_merge'] != 'both']

# 엑셀 파일에서 필요한 데이터 로드 (PA 데이터)
pa_invoice_data_pa = pd.read_excel(pa_invoice_file, sheet_name='CHALRTON CABINETRY CT', usecols=['Type', 'Date', 'Num', 'Debit'])
po_ct_data = pd.read_excel(po_ct_file, sheet_name='Charlton PA', usecols=['Type', 'Date', 'Num', 'Credit'])
merged_data_pa = pd.merge(pa_invoice_data_pa, po_ct_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_pa = merged_data_pa[merged_data_pa['_merge'] != 'both']

# 결과를 엑셀 파일로 저장
output_file = "/Users/mymac/Desktop/PA.xlsx"
with pd.ExcelWriter(output_file) as writer:
    unique_values_nj.to_excel(writer, sheet_name='pa2ny', index=False)
    unique_values_ct.to_excel(writer, sheet_name='pa2nj', index=False)
    unique_values_pa.to_excel(writer, sheet_name='pa2ct', index=False)
