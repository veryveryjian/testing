import pandas as pd
import numpy as np

# 엑셀 파일 경로
pa_invoice_file = "C:\\Users\\charlton\\Desktop\\pa_invoice_data.xlsx"
po_ny_file = "C:\\Users\\charlton\\Desktop\\po_ny_data.xlsx"
po_nj_file = "C:\\Users\\charlton\\Desktop\\po_nj_data.xlsx"
po_ct_file = "C:\\Users\\charlton\\Desktop\\po_ct_data.xlsx"  # 수정된 부분
po_mi_file = "C:\\Users\\charlton\\Desktop\\po_mi_data.xlsx"  #만들어여 돼 데이터가 없어서 일단 pass  인보이스만 가져오도록 우선 설정


# 'Chunghua PA' 시트에서 데이터 로드
po_ny_data1 = pd.read_excel(po_ny_file, sheet_name='Chunghua PA', usecols=['Type', 'Date', 'Num', 'Credit'])
# 'Charlton PA INC' 시트에서 데이터 로드
po_ny_data2 = pd.read_excel(po_ny_file, sheet_name='Charlton PA INC', usecols=['Type', 'Date', 'Num', 'Credit'])
# 두 시트의 데이터를 결합
po_ny_data = pd.concat([po_ny_data1, po_ny_data2], ignore_index=True)

#print(po_ny_data1)
#print(po_ny_data2)
#print(po_ny_data)


# 엑셀 파일에서 필요한 데이터 로드 (NJ 데이터)
pa_invoice_data_nj = pd.read_excel(pa_invoice_file, sheet_name='CHARLTON CABINETRY NY', usecols=['Type', 'Date', 'Num', 'Debit'])
merged_data_nj = pd.merge(pa_invoice_data_nj, po_ny_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_nj = merged_data_nj[merged_data_nj['_merge'] != 'both']

#######################################################

###새로 추가된 코드 Debit/Credit
merged_data2 = pd.merge(pa_invoice_data_nj, po_ny_data, on=['Type', 'Date', 'Num'], how='outer')
merged_data2['Debit/Credit'] = np.where(merged_data2['Debit'].notna(), merged_data2['Debit'], merged_data2['Credit'])
merged_data2.drop(columns=['Debit', 'Credit'], inplace=True)

#중복
# 각 Debit/Credit 값의 빈도 계산
value_counts = merged_data2['Debit/Credit'].value_counts()
merged_data2['Frequency'] = merged_data2['Debit/Credit'].map(value_counts)

#합계값
total_debit = pa_invoice_data_nj['Debit'].sum()
total_credit = po_ny_data['Credit'].sum()
# 차액 계산
balance = total_debit - total_credit
####
# 필터링하여 Frequency 값이 1인 행만 출력
frequency_one_rows = merged_data2[merged_data2['Frequency'] == 1]
# Type별로 Debit/Credit 총합계를 구함
totals_by_type = frequency_one_rows.groupby('Type')['Debit/Credit'].sum()
difference = totals_by_type['Invoice'] - totals_by_type['Purchase Order']
check = balance - difference

print("Debit 열의 총합계: ", total_debit)
print("Credit 열의 총합계: ", total_credit)
print("차액 (total_debit - total_credit): ", balance)
print(frequency_one_rows)
print(totals_by_type)
print("Difference (Invoice - Purchase Order):", difference)
print("=====================================================Finish==================================================================")


# 엑셀 파일에서 필요한 데이터 로드 (CT 데이터)
pa_invoice_data_ct = pd.read_excel(pa_invoice_file, sheet_name='CHARLTON CABINETRY NJ', usecols=['Type', 'Date', 'Num', 'Debit'])
po_nj_data = pd.read_excel(po_nj_file, sheet_name='Charlton Cabinetry - PA', usecols=['Type', 'Date', 'Num', 'Credit'])
merged_data_ct = pd.merge(pa_invoice_data_ct, po_nj_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_ct = merged_data_ct[merged_data_ct['_merge'] != 'both']

merged_data3 = pd.merge(pa_invoice_data_ct, po_nj_data, on=['Type', 'Date', 'Num'], how='outer')
merged_data3['Debit/Credit'] = np.where(merged_data3['Debit'].notna(), merged_data3['Debit'], merged_data3['Credit'])
merged_data3.drop(columns=['Debit', 'Credit'], inplace=True)
#print(merged_data3,"===merged_data3===")

#중복
# 각 Debit/Credit 값의 빈도 계산
value_counts_ct = merged_data3['Debit/Credit'].value_counts()
###########################################################################주의요망##################################
merged_data3['Frequency'] = merged_data3['Debit/Credit'].map(value_counts_ct)
#print("=====================123123123123========================")
#print(value_counts_ct, "value_counts_ct")
#print(merged_data3,"merged_data3")

#합계값
total_debit_ct = pa_invoice_data_ct['Debit'].sum()
total_credit_ct = po_nj_data['Credit'].sum()
print(total_debit_ct,"total_debit_ct")
print(total_credit_ct,"total_credit_ct")

# 차액 계산
balance_ct = total_debit_ct - total_credit_ct
print(balance_ct, "BLACNCE")

####
# 필터링하여 Frequency 값이 1인 행만 출력
frequency_one_rows_ct = merged_data3[merged_data3['Frequency'] == 1]
print(frequency_one_rows_ct)
# Type별로 Debit/Credit 총합계를 구함

totals_by_type_ct = frequency_one_rows_ct.groupby('Type')['Debit/Credit'].sum()
print(totals_by_type_ct)

difference_ct = totals_by_type_ct['Invoice'] - totals_by_type_ct['Purchase Order']
print(difference_ct)

check_ct = balance_ct - difference_ct
#print(check_ct)
print("=====================================================Finish==================================================================")


# 엑셀 파일에서 필요한 데이터 로드 (PA 데이터)
pa_invoice_data_pa = pd.read_excel(pa_invoice_file, sheet_name='CHALRTON CABINETRY CT', usecols=['Type', 'Date', 'Num', 'Debit'])
po_ct_data = pd.read_excel(po_ct_file, sheet_name='Charlton PA', usecols=['Type', 'Date', 'Num', 'Credit'])
merged_data_pa = pd.merge(pa_invoice_data_pa, po_ct_data, left_on=['Type', 'Date', 'Num', 'Debit'],
                          right_on=['Type', 'Date', 'Num', 'Credit'], how='outer', indicator=True)
unique_values_pa = merged_data_pa[merged_data_pa['_merge'] != 'both']


#####
merged_data4 = pd.merge(pa_invoice_data_pa, po_ct_data, on=['Type', 'Date', 'Num'], how='outer')
merged_data4['Debit/Credit'] = np.where(merged_data4['Debit'].notna(), merged_data4['Debit'], merged_data4['Credit'])
merged_data4.drop(columns=['Debit', 'Credit'], inplace=True)
#print(merged_data4,"===merged_data3===")

print(merged_data4)

#중복
# 각 Debit/Credit 값의 빈도 계산
value_counts_pa = merged_data4['Debit/Credit'].value_counts()
print(value_counts_pa,"test============================================================")
###########################################################################주의요망##################################
merged_data4['Frequency'] = merged_data4['Debit/Credit'].map(value_counts_pa)
#print("=====================123123123123========================")
#print(value_counts_pa)
#print(merged_data4)

#합계값
total_debit_pa = pa_invoice_data_pa['Debit'].sum()
total_credit_pa = po_ct_data['Credit'].sum()
print("total_debit_pa :",total_debit_pa)
print("total_credit_pa :",total_credit_pa)

# 차액 계산
balance_pa = total_debit_pa - total_credit_pa
print( "BLACNCE : ",balance_pa)
##################################################################################ErroR
##################################################################################NO data
####
# 필터링하여 Frequency 값이 1인 행만 출력
frequency_one_rows_pa = merged_data4[merged_data4['Frequency'] == 1]
print("frequency_one_rows_pa:",frequency_one_rows_pa)
# Type별로 Debit/Credit 총합계를 구함

totals_by_type_pa = frequency_one_rows_pa.groupby('Type')['Debit/Credit'].sum()
print("totals_by_type_pa :",totals_by_type_pa)

difference_pa = totals_by_type_pa['Invoice'] - totals_by_type_pa['Purchase Order']
print("difference_pa :",difference_pa)



#mi inv
ny_invoice_data_mi = pd.read_excel(pa_invoice_file, sheet_name='Miller Cabinetry Inc', usecols=['Type', 'Date', 'Num', 'Debit'])
unique_values_mi = ny_invoice_data_mi


# 결과를 엑셀 파일로 저장
#output_file = "C:\\Users\\charlton\\Desktop\\PA.xlsx"
#with pd.ExcelWriter(output_file) as writer:
#    unique_values_nj.to_excel(writer, sheet_name='pa2ny', index=False)
#    unique_values_ct.to_excel(writer, sheet_name='pa2nj', index=False)
#    unique_values_pa.to_excel(writer, sheet_name='pa2ct', index=False)
#    unique_values_mi.to_excel(writer, sheet_name='pa2mi', index=False)

