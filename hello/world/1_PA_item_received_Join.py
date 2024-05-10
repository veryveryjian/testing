import pandas as pd

# Define file paths
ny_invoice_file = "C:\\Users\\charlton\\Desktop\\pa_invoice_data.xlsx"
po_nj_file = "C:\\Users\\charlton\\Desktop\\Item_receive\\NJ_Received.xlsx"
po_ny_file = "C:\\Users\\charlton\\Desktop\\Item_receive\\NY_Received.xlsx"
po_pa_file = "C:\\Users\\charlton\\Desktop\\Item_receive\\CT_Received.xlsx"


# 엑셀 파일 로드
#xl = pd.ExcelFile(ny_invoice_file)

# 모든 시트 이름 출력
#print(xl.sheet_names)

# Load data from Excel files
# Convert 'Num' columns to string to ensure consistent data types for merging
ny_invoice_data_nj = pd.read_excel(ny_invoice_file, sheet_name='CHARLTON CABINETRY NY', usecols=['Type', 'Date', 'Num', 'Debit'])
ny_invoice_data_nj['Num'] = ny_invoice_data_nj['Num'].astype(str)

po_nj_data = pd.read_excel(po_nj_file, sheet_name='Charlton Cabinetry - PA', usecols=['Vendor', 'Date', 'Num', 'Amount'])
po_nj_data['Num'] = po_nj_data['Num'].astype(str)

# Perform similar conversions for CT and PA data
ny_invoice_data_ct = pd.read_excel(ny_invoice_file, sheet_name='CHARLTON CABINETRY NJ', usecols=['Type', 'Date', 'Num', 'Debit'])
ny_invoice_data_ct['Num'] = ny_invoice_data_ct['Num'].astype(str)

po_ct_data = pd.read_excel(po_ny_file, sheet_name='newPA', usecols=['Vendor', 'Date', 'Num', 'Amount'])
##################################################Chunghua PA  이것도 추가해야 돼
po_ct_data['Num'] = po_ct_data['Num'].astype(str)

ny_invoice_data_pa = pd.read_excel(ny_invoice_file, sheet_name='CHALRTON CABINETRY CT', usecols=['Type', 'Date', 'Num', 'Debit'])
ny_invoice_data_pa['Num'] = ny_invoice_data_pa['Num'].astype(str)

po_pa_data = pd.read_excel(po_pa_file, sheet_name='Chung Hua PA', usecols=['Vendor', 'Date', 'Num', 'Amount'])
################################################## CHUNG HUA CT 이것도 추가해야 돼
po_pa_data['Num'] = po_pa_data['Num'].astype(str)

# Merge data
merged_data_nj = pd.merge(ny_invoice_data_nj, po_nj_data, on=[ 'Date', 'Num'], how='outer', indicator=True)
unique_values_nj = merged_data_nj[merged_data_nj['_merge'] != 'both']

merged_data_ct = pd.merge(ny_invoice_data_ct, po_ct_data, on=[ 'Date', 'Num'], how='outer', indicator=True)
unique_values_ct = merged_data_ct[merged_data_ct['_merge'] != 'both']

merged_data_pa = pd.merge(ny_invoice_data_pa, po_pa_data, on=[ 'Date', 'Num'], how='outer', indicator=True)
unique_values_pa = merged_data_pa[merged_data_pa['_merge'] != 'both']

# Save results to Excel file
output_file = "C:\\Users\\charlton\\Desktop\\PA.xlsx"
with pd.ExcelWriter(output_file) as writer:
    unique_values_nj.to_excel(writer, sheet_name='pa2nj', index=False)
    unique_values_ct.to_excel(writer, sheet_name='pa2ny', index=False)
    unique_values_pa.to_excel(writer, sheet_name='pa2ct', index=False)


