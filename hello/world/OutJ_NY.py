import pandas as pd

# 엑셀 파일 경로
ny_invoice_file = "C:\\Users\\charlton\\Desktop\\ny_invoice_data.xlsx"
po_nj_file = "C:\\Users\\charlton\\Desktop\\Item_receive\\NJ_Received.xlsx"
po_ct_file = "C:\\Users\\charlton\\Desktop\\Item_receive\\CT_Received.xlsx"
po_pa_file = "C:\\Users\\charlton\\Desktop\\Item_receive\\PA_Received.xlsx"
po_mi_file = "C:\\Users\\charlton\\Desktop\\Item_receive\\NJ_Received.xlsx"


# 엑셀 파일에서 필요한 데이터 로드 (NJ 데이터)
ny_invoice_data_nj = pd.read_excel(ny_invoice_file, sheet_name='Chunghua NJ', usecols=['Type', 'Date', 'Num', 'Debit',])
po_nj_data = pd.read_excel(po_nj_file, sheet_name='Charlton Cabinetry - NY', usecols=['Vendor', 'Date', 'Num', 'Amount'])
merged_data_nj = pd.merge(ny_invoice_data_nj, po_nj_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Vendor', 'Date', 'Num', 'Amount'], how='outer', indicator=True)
unique_values_nj = merged_data_nj[merged_data_nj['_merge'] != 'both']

# 엑셀 파일에서 필요한 데이터 로드 (CT 데이터)
ny_invoice_data_ct = pd.read_excel(ny_invoice_file, sheet_name='Chunghua CT', usecols=['Type', 'Date', 'Num', 'Debit'])
po_ct_data = pd.read_excel(po_ct_file, sheet_name='CHARLTON CABINETRY NY', usecols=['Vendor', 'Date', 'Num', 'Amount'])
merged_data_ct = pd.merge(ny_invoice_data_ct, po_ct_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Vendor', 'Date', 'Num', 'Amount'], how='outer', indicator=True)
unique_values_ct = merged_data_ct[merged_data_ct['_merge'] != 'both']

# 엑셀 파일에서 필요한 데이터 로드 (PA 데이터)
ny_invoice_data_pa = pd.read_excel(ny_invoice_file, sheet_name='Charlton PA', usecols=['Type', 'Date', 'Num', 'Debit'])
po_pa_data = pd.read_excel(po_pa_file, sheet_name='CHARLTON NY', usecols=['Vendor', 'Date', 'Num', 'Amount'])  # 수정된 부분
merged_data_pa = pd.merge(ny_invoice_data_pa, po_pa_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_pa = merged_data_pa[merged_data_pa['_merge'] != 'both']


# 결과를 엑셀 파일로 저장
output_file = "C:\\Users\\charlton\\Desktop\\NY.xlsx"
with pd.ExcelWriter(output_file) as writer:
    unique_values_nj.to_excel(writer, sheet_name='ny2nj', index=False)
    unique_values_ct.to_excel(writer, sheet_name='ny2ct', index=False)
    unique_values_pa.to_excel(writer, sheet_name='ny2pa', index=False)


