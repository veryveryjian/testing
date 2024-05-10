import pandas as pd

# NJ CSV 파일 읽기
df_nj = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\nj.CSV", encoding='ISO-8859-1')
# NJ 필터링할 이름 목록
nj_vendor_names = ['Charlton Cabinetry - NY', 'Charlton Cabinetry - PA', 'Chunghua - CT', 'MILLER CABINETRY']

# df_nj 데이터에서 NJ 이름들에 해당하는 데이터만 필터링
NJ_filtered_data = df_nj[df_nj['Vendor'].isin(nj_vendor_names)]

# NJ 필터링된 데이터 출력
print("NJ 필터링된 데이터:")
print(NJ_filtered_data)

# NJ 필터링된 데이터를 엑셀 파일로 저장
with pd.ExcelWriter("C:\\Users\\charlton\\Desktop\\Item_receive\\NJ_Received.xlsx") as writer:
    ################################################################이름
    for vendor in nj_vendor_names:
        filtered_df = NJ_filtered_data[NJ_filtered_data['Vendor'] == vendor]
        filtered_df.to_excel(writer, sheet_name=vendor[:31], index=False)

print("NJ 엑셀 파일 저장 완료!")

# NY CSV 파일 읽기
df_ny = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\ny.CSV", encoding='ISO-8859-1')

# NY 필터링할 이름 목록
ny_vendor_names = ['CHUNG HUA CABINET CT INC', 'CHUNG HUA CABINET NJ INC', 'Charlton PA INC', 'Chunghua PA', 'Miller Cabinetry Inc']

# df_ny 데이터에서 NY 이름들에 해당하는 데이터만 필터링
NY_filtered_data = df_ny[df_ny['Vendor'].isin(ny_vendor_names)]

# NY 필터링된 데이터 출력
print("\nNY 필터링된 데이터:")
print(NY_filtered_data)

# NY 필터링된 데이터를 엑셀 파일로 저장
with pd.ExcelWriter("C:\\Users\\charlton\\Desktop\\Item_receive\\NY_Received.xlsx") as writer:
    for vendor in ny_vendor_names:
        filtered_df = NY_filtered_data[NY_filtered_data['Vendor'] == vendor]
        filtered_df.to_excel(writer, sheet_name=vendor[:31], index=False)

print("NY 엑셀 파일 저장 완료!")
