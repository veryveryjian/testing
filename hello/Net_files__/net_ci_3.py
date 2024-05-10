import os
import pandas as pd

# 경로 설정
paths = {
    "CI": r"\\NAS132B9F\1.Purchase order _订单&统计\1-8订单\1-8-22 越南三洋—出货单",
    "PI": r"\\NAS132B9F\1.Purchase order _订单&统计\1-8订单\1-8-23越南三洋订单\PI",
    "PO": r"\\NAS132B9F\1.Purchase order _订单&统计\1-8订单\1-8-23越南三洋订单\PO"
}


# 데이터프레임을 생성하고 엑셀로 저장하는 함수
def save_excel_files():
    for key, path in paths.items():
        # 해당 경로에 있는 파일들을 리스트로 가져오기
        files = os.listdir(path)
        # 엑셀 파일만 필터링하여 파일 이름 저장
        excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xls')]
        # 파일 이름을 데이터프레임에 저장
        df = pd.DataFrame(excel_files, columns=['File Name'])

        # 바탕화면 경로 설정 (윈도우 사용자의 경우)
        desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        # 데이터프레임을 엑셀 파일로 저장
        df.to_excel(os.path.join(desktop_path, f"{key}_files.xlsx"), index=False)


save_excel_files()
