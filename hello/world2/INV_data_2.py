import pandas as pd
print("NY 2 >>>")

# ny.CSV 파일의 경로
ny_path = "C:\\Users\\charlton\\Desktop\\INV_Monthly\\ny.CSV"

# 파일을 읽어와서 데이터프레임에 저장
df = pd.read_csv(ny_path)

# Excel 파일 생성을 위한 writer 객체
excel_writer = pd.ExcelWriter('C:\\Users\\charlton\\Desktop\\ny__.xlsx', engine='xlsxwriter')

# 'Vendor'별로 그룹화 및 처리
for Name, group in df.groupby('Name'):
    # 각 그룹의 'Amount' 합계 계산
    total_amount = group['Debit'].sum()

    # 합계 출력
    print(f"Name: {Name}, Total Amount: {total_amount}")

    # 각 그룹을 별도의 시트로 저장
    group.to_excel(excel_writer, sheet_name=Name, index=False)

    # 현재 시트에 합계 추가
    worksheet = excel_writer.sheets[Name]
    last_row = len(group) + 2
    worksheet.write(f'A{last_row}', 'Total Amount')
    worksheet.write(f'B{last_row}', total_amount)

# Excel 파일 저장 및 종료
excel_writer.close()
print("======================================================================================================================")
# ===========================================================================================================================
print("Nj 2 >>>")

import pandas as pd

# ny.CSV 파일의 경로
ny_path = "C:\\Users\\charlton\\Desktop\\INV_Monthly\\nj.CSV"

# 파일을 읽어와서 데이터프레임에 저장
df = pd.read_csv(ny_path)

# Excel 파일 생성을 위한 writer 객체
excel_writer = pd.ExcelWriter('C:\\Users\\charlton\\Desktop\\nj__.xlsx', engine='xlsxwriter')

# 'Vendor'별로 그룹화 및 처리
for Name, group in df.groupby('Name'):
    # 각 그룹의 'Amount' 합계 계산
    total_amount = group['Debit'].sum()

    # 합계 출력
    print(f"Name: {Name}, Total Amount: {total_amount}")

    # 각 그룹을 별도의 시트로 저장
    group.to_excel(excel_writer, sheet_name=Name, index=False)

    # 현재 시트에 합계 추가
    worksheet = excel_writer.sheets[Name]
    last_row = len(group) + 2
    worksheet.write(f'A{last_row}', 'Total Amount')
    worksheet.write(f'B{last_row}', total_amount)

# Excel 파일 저장 및 종료
excel_writer.close()
print("======================================================================================================================")
# ==========================================================================================================================
print("CT 2 >>>")

import pandas as pd

# ny.CSV 파일의 경로
ny_path = "C:\\Users\\charlton\\Desktop\\INV_Monthly\\ct.CSV"

# 파일을 읽어와서 데이터프레임에 저장
df = pd.read_csv(ny_path)

# Excel 파일 생성을 위한 writer 객체
excel_writer = pd.ExcelWriter('C:\\Users\\charlton\\Desktop\\ct__.xlsx', engine='xlsxwriter')

# 'Vendor'별로 그룹화 및 처리
for Name, group in df.groupby('Name'):
    # 각 그룹의 'Amount' 합계 계산
    total_amount = group['Debit'].sum()

    # 합계 출력
    print(f"Name: {Name}, Total Amount: {total_amount}")

    # 각 그룹을 별도의 시트로 저장
    group.to_excel(excel_writer, sheet_name=Name, index=False)

    # 현재 시트에 합계 추가
    worksheet = excel_writer.sheets[Name]
    last_row = len(group) + 2
    worksheet.write(f'A{last_row}', 'Total Amount')
    worksheet.write(f'B{last_row}', total_amount)

# Excel 파일 저장 및 종료
excel_writer.close()
print("======================================================================================================================")
# ===========================================================================================================================
print("PA 2 >>>")

import pandas as pd

# ny.CSV 파일의 경로
ny_path = "C:\\Users\\charlton\\Desktop\\INV_Monthly\\pa.CSV"

# 파일을 읽어와서 데이터프레임에 저장
df = pd.read_csv(ny_path)

# Excel 파일 생성을 위한 writer 객체
excel_writer = pd.ExcelWriter('C:\\Users\\charlton\\Desktop\\pa__.xlsx', engine='xlsxwriter')

# 'Vendor'별로 그룹화 및 처리
for Name, group in df.groupby('Name'):
    # 각 그룹의 'Amount' 합계 계산
    total_amount = group['Debit'].sum()

    # 합계 출력
    print(f"Name: {Name}, Total Amount: {total_amount}")

    # 각 그룹을 별도의 시트로 저장
    group.to_excel(excel_writer, sheet_name=Name, index=False)

    # 현재 시트에 합계 추가
    worksheet = excel_writer.sheets[Name]
    last_row = len(group) + 2
    worksheet.write(f'A{last_row}', 'Total Amount')
    worksheet.write(f'B{last_row}', total_amount)

# Excel 파일 저장 및 종료
excel_writer.close()
print("======================================================================================================================")
# ===========================================================================================================================

