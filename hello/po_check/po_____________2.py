import pandas as pd

# 네트워크 드라이브 또는 공유 폴더의 엑셀 파일 경로 설정
excel_file = r'\\NAS132B9F\1.Purchase order _订单&统计\1-11工厂价格\三洋价格\三洋越南价格\三洋工厂价格越南.xlsx'

try:
    # 엑셀 파일의 모든 시트 이름 불러오기
    xl = pd.ExcelFile(excel_file)
    # 필요한 시트 이름 설정
    needed_sheets = ['MS', 'SG', 'SE', 'CW2', 'E2']
    # 각 시트에서 필요한 열 데이터를 담을 빈 데이터프레임 생성
    price = pd.DataFrame()

    for sheet_name in needed_sheets:
        # 현재 시트의 데이터 불러오기
        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        # 시트 이름에 따라 적절한 열 선택
        if sheet_name == 'E2':
            column_name = '单价'
        else:
            column_name = '价格'

        # 선택한 열 데이터를 price 데이터프레임에 추가
        # 새로운 열 이름을 시트 이름으로 설정하여 구분 가능하게 함
        price[sheet_name] = df[column_name]

    # 결과 출력
    print("Collected Prices DataFrame:")
    print(price)

except FileNotFoundError:
    print("파일을 찾을 수 없습니다. 파일 경로를 확인하세요.")
