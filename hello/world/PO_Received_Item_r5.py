import pandas as pd

# CSV 파일 읽기
df_ny = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\ny.CSV", encoding='ISO-8859-1')
df_nj = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\nj.CSV", encoding='ISO-8859-1')
df_ct = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\ct.CSV", encoding='ISO-8859-1')
df_pa = pd.read_csv("C:\\Users\\charlton\\Desktop\\Item_receive\\pa.CSV", encoding='ISO-8859-1')

# 각 주별 필터링할 이름 목록
ny_vendor_names = ['CHUNG HUA CABINET NJ INC', 'CHUNG HUA CABINET CT INC', 'Charlton PA INC', 'Chunghua PA', 'Miller Cabinetry Inc']
nj_vendor_names = ['Charlton Cabinetry - NY', 'Chunghua - CT', 'Charlton Cabinetry - PA', 'MILLER CABINETRY']
ct_vendor_names = ['CHARLTON CABINETRY NY', 'Chung Hua Cabinet NJ Inc', 'Chung Hua PA', 'Miller Cabinetry']
pa_vendor_names = ['CHARLTON NY', 'CHARLTON NJ', 'CHARLTON CT', 'CHUNG HUA CT']

# NY 데이터 필터링 및 저장
with pd.ExcelWriter("C:\\Users\\charlton\\Desktop\\Item_receive\\NY_Received.xlsx") as writer:
    newPA = pd.DataFrame()
    for vendor in ny_vendor_names:
        filtered_df = df_ny[df_ny['Vendor'] == vendor]
        if vendor in ['Charlton PA INC', 'Chunghua PA']:
            newPA = pd.concat([newPA, filtered_df])
        else:
            filtered_df.to_excel(writer, sheet_name=vendor[:31], index=False)
    newPA.to_excel(writer, sheet_name='newPA', index=False)
print("NY 엑셀 파일 저장 완료!")

# NJ 데이터 필터링 및 저장
with pd.ExcelWriter("C:\\Users\\charlton\\Desktop\\Item_receive\\NJ_Received.xlsx") as writer:
    for vendor in nj_vendor_names:
        filtered_df = df_nj[df_nj['Vendor'] == vendor]
        filtered_df.to_excel(writer, sheet_name=vendor[:31], index=False)
print("NJ 엑셀 파일 저장 완료!")

# CT 데이터 필터링 및 저장
with pd.ExcelWriter("C:\\Users\\charlton\\Desktop\\Item_receive\\CT_Received.xlsx") as writer:
    for vendor in ct_vendor_names:
        filtered_df = df_ct[df_ct['Vendor'] == vendor]
        filtered_df.to_excel(writer, sheet_name=vendor[:31], index=False)
print("CT 엑셀 파일 저장 완료!")

# PA 데이터 필터링 및 저장
with pd.ExcelWriter("C:\\Users\\charlton\\Desktop\\Item_receive\\PA_Received.xlsx") as writer:
    newPA = pd.DataFrame()
    for vendor in pa_vendor_names:
        filtered_df = df_pa[df_pa['Vendor'] == vendor]
        if vendor in ['CHARLTON CT', 'CHUNG HUA CT']:
            newPA = pd.concat([newPA, filtered_df])
        else:
            filtered_df.to_excel(writer, sheet_name=vendor[:31], index=False)
    newPA.to_excel(writer, sheet_name='newCT', index=False)
print("PA 엑셀 파일 저장 완료!")
