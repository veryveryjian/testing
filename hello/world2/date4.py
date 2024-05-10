import pandas as pd

# 이름 변경을 위한 사전 (dictionary) 정의
name_changes = {
    "CHUNG HUA CABINET CT INC": "CT2NY",
    "CHUNG HUA CABINET NJ INC": "NJ2NY",
    "Charlton PA INC": "PA2NY",
    "Miller Cabinetry Inc": "MI2NY",
    # 다른 이름 매핑도 이곳에 추가하세요...
}

def process_file(csv_path, excel_path):
    # 파일을 읽어와서 데이터프레임에 저장
    df = pd.read_csv(csv_path)

    # Excel 파일 생성을 위한 writer 객체
    excel_writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')

    # 'Vendor' 또는 'Name'별로 그룹화 및 처리
    for name, group in df.groupby('Vendor'):  # 'Vendor' 또는 'Name'에 따라 변경
        # "Jan 24" 제외
        if name != "Jan 24":
            # 이름 변경
            display_name = name_changes.get(name, name)

            # 각 그룹의 'Amount' 합계 계산
            total_amount = group['Amount'].sum()

            # 합계 출력
            print(f"{display_name}, Total Amount: {total_amount}")

            # 각 그룹을 별도의 시트로 저장
            group.to_excel(excel_writer, sheet_name=display_name, index=False)

            # 현재 시트에 합계 추가
            worksheet = excel_writer.sheets[display_name]
            last_row = len(group) + 2
            worksheet.write(f'A{last_row}', 'Total Amount')
            worksheet.write(f'B{last_row}', total_amount)

    # Excel 파일 저장 및 종료
    excel_writer.close()
    print("==========================================================================================================")

# 각 파일에 대해 process_file 함수 호출
process_file("C:\\Users\\charlton\\Desktop\\IR_Monthly\\ny.CSV", 'C:\\Users\\charlton\\Desktop\\ny_.xlsx')
# 나머지 파일들에 대해서도 비슷한 방식으로 호출
# process_file("파일경로", "엑셀경로")
