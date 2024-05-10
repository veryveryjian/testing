import pandas as pd

# CSV 파일 읽기
df_ny = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\ny.CSV", encoding='ISO-8859-1')
df_nj = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\nj.CSV", encoding='ISO-8859-1')
df_ct = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\ct.CSV", encoding='ISO-8859-1')
df_pa = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\pa.CSV", encoding='ISO-8859-1')



vendor_filter= {
    'NY': ['CHUNG HUA CABINET CT INC', 'CHUNG HUA CABINET NJ INC', 'Charlton PA INC', 'Chunghua PA', 'Miller Cabinetry Inc'],
    'NJ': ['Charlton Cabinetry - NY', 'Charlton Cabinetry - PA','Chunghua - CT', 'MILLER CABINETRY' ],
    'CT': ['Miller Cabinetry', 'Chung Hua PA', 'Chung Hua Cabinet NJ Inc', 'Charlton PA', 'CHARLTON CABINETRY NY'],
    'PA': ['CHARLTON CABINETRY NY', 'CHARLTON CABINETRY NJ', 'CHALRTON CABINETRY CT', 'Miller Cabinetry Inc'],
    'MI': ['CHARLTON CT','CHARLTON NJ','CHARLTON NY','CHUNG HUA CT']
}


# 엑셀 파일로 저장하기 위한 ExcelWriter 생성
with pd.ExcelWriter("C:\\Users\\charlton\\Desktop\\Item_receive\\filtered_data111.xlsx") as writer:
    for vendor in vendor_filter:
        # 각 이름에 대해 필터링된 데이터프레임 생성
        filtered_df = df_ny[df_ny['Vendor'] == vendor]

        # 필터링된 데이터프레임을 별도의 시트로 저장
        filtered_df.to_excel(writer, sheet_name=vendor[:31], index=False)  # 시트 이름은 최대 31자까지 가능

print("엑셀 파일 저장 완료!")
