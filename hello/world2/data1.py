import pandas as pd

# ny.CSV 파일의 경로
ny_path = "C:\\Users\\charlton\\Desktop\\IR_Monthly\\ny.CSV"
# 파일을 읽어와서 데이터프레임에 저장
df = pd.read_csv(ny_path)
# Excel 파일 생성을 위한 writer 객체
excel_writer = pd.ExcelWriter('C:\\Users\\charlton\\Desktop\\Vendor_Sheets.xlsx', engine='xlsxwriter')
# 'Vendor'별로 그룹화
for vendor, group in df.groupby('Name'):
    # 각 그룹을 별도의 시트로 저장
    group.to_excel(excel_writer, sheet_name=vendor, index=False)
# Excel 파일 저장 및 종료
excel_writer.close()
