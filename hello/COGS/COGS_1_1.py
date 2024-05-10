import os
import pandas as pd

# 경로 설정
directory = "C:/Users/charlton/Desktop/COGS"

# 해당 경로의 모든 파일을 순회하며 CSV 파일을 찾고 출력
for filename in os.listdir(directory):
    if filename.endswith(".csv") or filename.endswith(".CSV"):
        # 파일의 전체 경로
        file_path = os.path.join(directory, filename)

        # CSV 파일 읽기
        df = pd.read_csv(file_path)

        # 데이터 출력
        print(f"파일 이름: {filename}")
        print(df)
        print("\n")
