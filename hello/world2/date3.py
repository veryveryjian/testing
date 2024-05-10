import pandas as pd

# ny.CSV 파일의 경로
ny_path = "C:\\Users\\charlton\\Desktop\\IR_Monthly\\ny.CSV"

# 파일을 읽어와서 데이터프레임에 저장
df = pd.read_csv(ny_path)

# Excel 파일 생성을 위한 writer 객체
excel_writer = pd.ExcelWriter('C:\\Users\\charlton\\Desktop\\ny_.xlsx', engine='xlsxwriter')

# 'Vendor'별로 그룹화 및 처리
for vendor, group in df.groupby('Vendor'):
    # 각 그룹의 'Amount' 합계 계산
    total_amount = group['Amount'].sum()

    # 합계 출력
    print(f"Vendor: {vendor}, Total Amount: {total_amount}")


    # 각 그룹을 별도의 시트로 저장
    group.to_excel(excel_writer, sheet_name=vendor, index=False)

    # 현재 시트에 합계 추가
    worksheet = excel_writer.sheets[vendor]
    last_row = len(group) + 2
    worksheet.write(f'A{last_row}', 'Total Amount')
    worksheet.write(f'B{last_row}', total_amount)

# Excel 파일 저장 및 종료
excel_writer.close()
print("==========================================================================================================")

# ===========================================================================================================================

import pandas as pd

# ny.CSV 파일의 경로
ny_path = "C:\\Users\\charlton\\Desktop\\IR_Monthly\\nj.CSV"

# 파일을 읽어와서 데이터프레임에 저장
df = pd.read_csv(ny_path)

# Excel 파일 생성을 위한 writer 객체
excel_writer = pd.ExcelWriter('C:\\Users\\charlton\\Desktop\\nj_.xlsx', engine='xlsxwriter')

# 'Vendor'별로 그룹화 및 처리
for Name, group in df.groupby('Name'):
    # 각 그룹의 'Amount' 합계 계산
    total_amount = group['Amount'].sum()

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
print("==========================================================================================================")

# ===========================================================================================================================

import pandas as pd

# ny.CSV 파일의 경로
ny_path = "C:\\Users\\charlton\\Desktop\\IR_Monthly\\ct.CSV"

# 파일을 읽어와서 데이터프레임에 저장
df = pd.read_csv(ny_path)

# Excel 파일 생성을 위한 writer 객체
excel_writer = pd.ExcelWriter('C:\\Users\\charlton\\Desktop\\ct_.xlsx', engine='xlsxwriter')

# 'Vendor'별로 그룹화 및 처리
for Name, group in df.groupby('Name'):
    # 각 그룹의 'Amount' 합계 계산
    total_amount = group['Amount'].sum()

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
print("==========================================================================================================")

# ===========================================================================================================================


import pandas as pd

# ny.CSV 파일의 경로
ny_path = "C:\\Users\\charlton\\Desktop\\IR_Monthly\\pa.CSV"

# 파일을 읽어와서 데이터프레임에 저장
df = pd.read_csv(ny_path)

# Excel 파일 생성을 위한 writer 객체
excel_writer = pd.ExcelWriter('C:\\Users\\charlton\\Desktop\\pa_.xlsx', engine='xlsxwriter')

# 'Vendor'별로 그룹화 및 처리
for Name, group in df.groupby('Name'):
    # 각 그룹의 'Amount' 합계 계산
    total_amount = group['Amount'].sum()

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
print("==========================================================================================================")

# ===========================================================================================================================