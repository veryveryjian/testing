import pandas as pd
import matplotlib.pyplot as plt

# CSV 파일 경로 (사용자 환경에서 경로를 확인해주세요)
file_path = 'C:\\Users\\charlton\\Desktop\\sale_analysis2.CSV'

# 데이터 로드
data = pd.read_csv(file_path)

# 데이터 출력
print(data)

# 그래프 스타일 설정
plt.style.use('ggplot')

# 새로운 figure 및 axis 생성
fig, ax1 = plt.subplots(figsize=(14, 7))

# 바 차트로 'Sum of Qty' 표시
ax1.bar(data['Cabinets sub-code'], data['Sum of Qty'], color='b', label='Sum of Qty', alpha=0.6)

# y축 레이블, 틱 및 틱 레이블 색상 설정
ax1.set_xlabel('Cabinets Sub-Code', fontsize=12)
ax1.set_ylabel('Sum of Qty', color='b', fontsize=12)
ax1.tick_params('y', colors='b')

# 두 번째 y축을 위해 x축 공유
ax2 = ax1.twinx()

# 라인 차트로 'Sum of CBM' 표시
ax2.plot(data['Cabinets sub-code'], data['Sum of CBM*Qty'], color='r', label='Sum of CBM*Qty', marker='o')
ax2.set_ylabel('Sum of CBM', color='r', fontsize=12)
ax2.tick_params('y', colors='r')

# 타이틀 및 레전드 설정
plt.title('Cabinet Sales Analysis', fontsize=16)
ax1.legend(loc='upper left')
ax2.legend(loc='upper right')

# 그래프 표시
plt.show()
