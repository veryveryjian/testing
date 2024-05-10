
#    'NJ': ['Charlton Cabinetry - NY', 'Charlton Cabinetry - PA','Chunghua - CT', 'MILLER CABINETRY' ],
#    'CT': ['Miller Cabinetry', 'Chung Hua PA', 'Chung Hua Cabinet NJ Inc', 'Charlton PA', 'CHARLTON CABINETRY NY'],
#    'PA': ['CHARLTON CABINETRY NY', 'CHARLTON CABINETRY NJ', 'CHALRTON CABINETRY CT', 'Miller Cabinetry Inc'],
#    'MI': ['CHARLTON CT','CHARLTON NJ','CHARLTON NY','CHUNG HUA CT']


# 엑셀 파일로 저장하기 위한 ExcelWriter 생성

import pandas as pd

# CSV 파일 읽기
df_nj = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\nj.CSV", encoding='ISO-8859-1')

# 필터링할 이름 목록
vendor_names = ['Charlton Cabinetry - NY', 'Charlton Cabinetry - PA','Chunghua - CT', 'MILLER CABINETRY']



# df_ny 데이터에서 위의 이름들에 해당하는 데이터만 필터링
NY_filtered_data = df_nj[df_nj['Vendor'].isin(vendor_names)]

# 필터링된 데이터 출력
print(NY_filtered_data)

# 엑셀 파일로 저장하기 위한 ExcelWriter 생성
with pd.ExcelWriter("C:\\Users\\charlton\\Desktop\\Item_receive\\NJ_Received.xlsx") as writer:
    ################################################################이름
    for vendor in vendor_names:
        # 각 이름에 대해 필터링된 데이터프레임 생성
        filtered_df = NY_filtered_data[NY_filtered_data['Vendor'] == vendor]

        # 필터링된 데이터프레임을 별도의 시트로 저장
        filtered_df.to_excel(writer, sheet_name=vendor[:31], index=False)  # 시트 이름은 최대 31자까지 가능

print("엑셀 파일 저장 완료!")
