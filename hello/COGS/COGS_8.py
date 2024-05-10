import pandas as pd

# 각 주에 대한 CSV 파일을 데이터프레임으로 읽어오기
ny = pd.read_csv("C:\\Users\\charlton\\Desktop\\COGS\\NY.CSV", encoding='ISO-8859-1')
nj = pd.read_csv("C:\\Users\\charlton\\Desktop\\COGS\\NJ.CSV", encoding='ISO-8859-1')
ct = pd.read_csv("C:\\Users\\charlton\\Desktop\\COGS\\CT.CSV", encoding='ISO-8859-1')
pa = pd.read_csv("C:\\Users\\charlton\\Desktop\\COGS\\PA.CSV", encoding='ISO-8859-1')

# NY 데이터프레임 출력
print("New York (NY) Data:")
print(ny.head())
print("\n")  # 줄바꿈을 위한 코드

# NJ 데이터프레임 출력
print("New Jersey (NJ) Data:")
print(nj.head())
print("\n")

# CT 데이터프레임 출력
print("Connecticut (CT) Data:")
print(ct.head())
print("\n")

# PA 데이터프레임 출력
print("Pennsylvania (PA) Data:")
print(pa.head())
print("\n")
