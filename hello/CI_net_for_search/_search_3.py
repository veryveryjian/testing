import os
import pandas as pd

def create_excel_with_ids(folder_path, save_path):
    folder_path = os.path.normpath(folder_path)  # 경로 정규화
    files_list = []

    # 지정된 폴더에서 엑셀 파일 목록 생성
    for file in os.listdir(folder_path):
        if file.endswith(('.xls', '.xlsx')) and not file.startswith('~$'):
            files_list.append(file)

    # 파일 목록을 데이터프레임으로 변환
    df = pd.DataFrame(files_list, columns=['FileName'])

    # 'ID' 컬럼 추가 및 'ID'를 첫 번째 컬럼으로 설정
    df.insert(0, 'ID', df.index + 1)  # 인덱스는 0부터 시작하므로 1을 더함

    # 엑셀 파일로 저장
    output_path = os.path.join(save_path, 'FileList_with_IDs.xlsx')
    df.to_excel(output_path, index=False)  # 인덱스는 엑셀 파일에 포함하지 않음

    print(f"File saved successfully at {output_path}")

# 사용 예시
folder_path = r'\\NAS132B9F\1.Purchase order _订单&统计\1-8订单\1-8-22 越南三洋—出货单'
desktop_path = r'C:\Users\charlton\Desktop'
create_excel_with_ids(folder_path, desktop_path)


