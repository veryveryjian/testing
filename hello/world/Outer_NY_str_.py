import pandas as pd

# 엑셀 파일 경로
ny_invoice_file = "C:\\Users\\charlton\\Desktop\\ny_invoice_data.xlsx"
po_nj_file = "C:\\Users\\charlton\\Desktop\\po_nj_data.xlsx"
po_ct_file = "C:\\Users\\charlton\\Desktop\\po_ct_data.xlsx"
po_pa_file = "C:\\Users\\charlton\\Desktop\\po_pa_data.xlsx"
po_mi_file = "C:\\Users\\charlton\\Desktop\\po_mi_data.xlsx"  # 데이터가 없으므로 일단 pass

# 엑셀 파일에서 필요한 데이터 로드
ny_invoice_data_nj = pd.read_excel(ny_invoice_file, sheet_name='Chunghua NJ', usecols=['Type', 'Date', 'Num', 'Debit'])
po_nj_data = pd.read_excel(po_nj_file, sheet_name='Charlton Cabinetry - NY', usecols=['Type', 'Date', 'Num', 'Credit', 'Memo'])

# 데이터 타입을 일치시키기 위해 'Num' 열을 문자열로 변환
ny_invoice_data_nj['Num'] = ny_invoice_data_nj['Num'].astype(str)
po_nj_data['Num'] = po_nj_data['Num'].astype(str)

# 병합 실행
merged_data_nj = pd.merge(ny_invoice_data_nj, po_nj_data, on=['Type', 'Date', 'Num'], how='outer', indicator=True)
unique_values_nj = merged_data_nj[merged_data_nj['_merge'] != 'both']

# CT 데이터
ny_invoice_data_ct = pd.read_excel(ny_invoice_file, sheet_name='Chunghua CT', usecols=['Type', 'Date', 'Num', 'Debit'])
po_ct_data = pd.read_excel(po_ct_file, sheet_name='CHARLTON CABINETRY NY', usecols=['Type', 'Date', 'Num', 'Credit', 'Memo'])

# 데이터 타입을 일치시키기 위해 'Num' 열을 문자열로 변환
ny_invoice_data_ct['Num'] = ny_invoice_data_ct['Num'].astype(str)
po_ct_data['Num'] = po_ct_data['Num'].astype(str)

# 병합 실행
merged_data_ct = pd.merge(ny_invoice_data_ct, po_ct_data, on=['Type', 'Date', 'Num'], how='outer', indicator=True)
unique_values_ct = merged_data_ct[merged_data_ct['_merge'] != 'both']

# PA 데이터
ny_invoice_data_pa = pd.read_excel(ny_invoice_file, sheet_name='Charlton PA', usecols=['Type', 'Date', 'Num', 'Debit'])
po_pa_data = pd.read_excel(po_pa_file, sheet_name='CHARLTON NY', usecols=['Type', 'Date', 'Num', 'Credit', 'Memo'])

# 데이터 타입을 일치시키기 위해 'Num' 열을 문자열로 변환
ny_invoice_data_pa['Num'] = ny_invoice_data_pa['Num'].astype(str)
po_pa_data['Num'] = po_pa_data['Num'].astype(str)

# 병합 실행
merged_data_pa = pd.merge(ny_invoice_data_pa, po_pa_data, on=['Type', 'Date', 'Num'], how='outer', indicator=True)
unique_values_pa = merged_data_pa[merged_data_pa['_merge'] != 'both']

# MI 데이터
ny_invoice_data_mi = pd.read_excel(ny_invoice_file, sheet_name='Miller Cabinetry', usecols=['Type', 'Date', 'Num', 'Debit'])
# po_mi_data 없이 'ny_invoice_data_mi'만 사용
unique_values_mi = ny_invoice_data_mi

# 결과를 엑셀 파일로 저장
output_file = "C:\\Users\\charlton\\Desktop\\NY.xlsx"
with pd.ExcelWriter(output_file) as writer:
    unique_values_nj.to_excel(writer, sheet_name='ny2nj', index=False)
    unique_values_ct.to_excel(writer, sheet_name='ny2ct', index=False)
    unique_values_pa.to_excel(writer, sheet_name='ny2pa', index=False)
    unique_values_mi.to_excel(writer, sheet_name='ny2mi', index=False)
