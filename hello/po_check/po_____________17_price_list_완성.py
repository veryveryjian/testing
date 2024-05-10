import pandas as pd
import os

# 네트워크 드라이브 또는 공유 폴더의 엑셀 파일 경로 설정
excel_file = r'\\NAS132B9F\1.Purchase order _订单&统计\1-11工厂价格\三洋价格\三洋越南价格\三洋工厂价格越南.xlsx'

# 최종 Excel 파일을 저장할 경로
file_path = 'C:/Users/charlton/Desktop/p.xlsx'

try:
    # 엑셀 파일의 모든 시트 이름 불러오기
    xl = pd.ExcelFile(excel_file)
    needed_sheets = ['MS', 'SG', 'SE', 'CW2', 'E2']

    # 최종 데이터를 담을 리스트
    final_data = []

    for sheet_name in needed_sheets:
        # 현재 시트의 데이터 불러오기
        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        # 시트 이름에 따라 적절한 열 선택
        if sheet_name == 'E2':
            price_column = '单价'
        else:
            price_column = '价格'

        # ITEM 열, 가격 열, 시트 이름을 포함하는 데이터프레임 생성
        temp_df = pd.DataFrame({
            'ITEM': df.iloc[:, 0],  # 첫 번째 열이 ITEM 열
            'Price': df[price_column],
            'Sheet Name': sheet_name  # 데이터가 어느 시트에서 왔는지
        })

        # 최종 데이터 리스트에 추가
        final_data.append(temp_df)

    # 리스트를 하나의 데이터프레임으로 결합
    final_df = pd.concat(final_data, ignore_index=True)

    # 결과 DataFrame을 Excel 파일로 저장
    final_df.to_excel(file_path, index=False)

    print(f"파일이 저장되었습니다: {file_path}")

except FileNotFoundError:
    print("파일을 찾을 수 없습니다. 파일 경로를 확인하세요.")



