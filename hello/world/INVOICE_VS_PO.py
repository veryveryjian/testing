import pandas as pd
#5306.82
# 파일 읽기
po = "C:\\Users\\charlton\\Desktop\\inv_po_compare\\4109.CSV"
po = pd.read_csv(po, encoding='ISO-8859-1')

# Num 열 처리: NaN 값이 있으면 그대로 두고, 나머지는 정수로 변환 후 문자열로 변경
po['Num'] = po['Num'].apply(lambda x: str(int(x)) if pd.notna(x) else x)

# 출력 옵션 설정: 최대 행 수와 최대 열 수를 None으로 설정하여 모든 행과 열을 출력하도록 함
#pd.set_option('display.max_rows', None)
#pd.set_option('display.max_columns', None)

# Num이 '4109'이고 Account가 'Inventory Asset'인 모든 행 필터링 및 출력
#filtered_rows = po[(po['Num'] == '4109') & (po['Account'] == 'Inventory Asset')]
filtered_rows = po[(po['Num'] == '4109') & (po['Account'] == 'Inventory Asset')]
print(filtered_rows)

# Debit 열의 합계 계산 및 출력
debit_sum = filtered_rows['Debit'].sum()
print("\nDebit 열의 합계:", debit_sum)

#output_path = "C:\\Users\\charlton\\Desktop\\invpo.xlsx"
#filtered_rows.to_excel(output_path, index=False)



####



import pandas as pd
#5158.67
# 파일 읽기
inv_path = "C:\\Users\\charlton\\Desktop\\inv_po_compare\\30867.CSV"
inv = pd.read_csv(inv_path, encoding='ISO-8859-1')

# Num 열 처리: NaN 값이 있으면 그대로 두고, 나머지는 정수로 변환 후 문자열로 변경
inv['Num'] = inv['Num'].apply(lambda x: str(int(x)) if pd.notna(x) else x)

# 출력 옵션 설정: 최대 행 수와 최대 열 수를 None으로 설정하여 모든 행과 열을 출력하도록 함
#pd.set_option('display.max_rows', None)
#pd.set_option('display.max_columns', None)

# Num이 '4109'이고 Account가 'Inventory Asset'인 모든 행 필터링 및 출력
filtered_rows = inv[(inv['Num'] == '30867') & (inv['Account'] == 'Inventory Asset')]
print(filtered_rows)

# Debit 열의 합계 계산 및 출력
credit_sum = filtered_rows['Credit'].sum()
print("\nDebit 열의 합계:", credit_sum)

#output_path = "C:\\Users\\charlton\\Desktop\\invpo.xlsx"
#filtered_rows.to_excel(output_path, index=False)
