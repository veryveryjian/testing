import os

# 디렉토리 경로 설정
directory_path = r"C:\Users\charlton\Desktop\pos2"

# 해당 디렉토리에 있는 모든 파일 및 폴더의 이름을 출력
for filename in os.listdir(directory_path):
    print(filename)


