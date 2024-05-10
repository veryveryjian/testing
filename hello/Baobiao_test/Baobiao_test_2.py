import openpyxl

# 엑셀 파일 경로 지정
file_path = r'C:\Users\charlton\Desktop\9-3 CT提货出货入库组装流程单 CT Pickup, Shipment and Warehouse Assembly Process Sheet.xlsx'

# 엑셀 파일 로드
workbook = openpyxl.load_workbook(file_path)

# 특정 시트 할당
assembly_sheet = workbook["组装安排-CT-Assembly"]
delivery_sheet = workbook["送货安排-CT Delivery"]

# 시트 데이터 출력 함수
def print_sheet_data(sheet):
    print(f"시트 이름: {sheet.title}\n")
    for row in sheet.iter_rows(min_row=1, max_row=10, values_only=True):  # 첫 10행만 출력하도록 제한
        print(row)
    print("\n" + "-"*40 + "\n")  # 시트 간 구분을 위해 줄 추가

# 각 시트의 데이터 출력
print_sheet_data(assembly_sheet)
print_sheet_data(delivery_sheet)
