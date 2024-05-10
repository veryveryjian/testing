import pandas as pd

# 파일 경로 정의
ny_test_file_path = 'C:/Users/charlton/Desktop/409-MS-JIN.xlsx'

# pandas를 사용하여 15행을 제외한 데이터 읽기
df = pd.read_excel(ny_test_file_path, skiprows=15)

# 읽어온 데이터 출력
print(df)
