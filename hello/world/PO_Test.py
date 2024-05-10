import pandas as pd

# CSV 파일 읽기
df_ny = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku_data 2\\ny_po.CSV", encoding='ISO-8859-1')
df_nj = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku_data 2\\nj_po.CSV", encoding='ISO-8859-1')
df_ct = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku_data 2\\ct_po.CSV", encoding='ISO-8859-1')
df_pa = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku_data 2\\pa_po.CSV", encoding='ISO-8859-1')
df_mi = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku_data 2\\mi_po.CSV", encoding='ISO-8859-1')


# 각 데이터프레임에 대해 'Num' 컬럼으로 오름차순 정렬
df_ny.sort_values(by='Num', inplace=True)
df_nj.sort_values(by='Num', inplace=True)
df_ct.sort_values(by='Num', inplace=True)
df_pa.sort_values(by='Num', inplace=True)
df_mi.sort_values(by='Num', inplace=True)


# 각 주별로 데이터프레임 처리
states = {'ny': df_ny, 'nj': df_nj, 'ct': df_ct, 'pa': df_pa, 'mi':df_mi}
for state, df in states.items():
    # Excel 파일 작성자 생성
    writer = pd.ExcelWriter(f'/Users\\charlton\\Desktop\\po_{state}_data.xlsx', engine='xlsxwriter')

    # 'Name' 컬럼의 고유한 값들에 대해 반복
    for name in df['Name'].dropna().unique():
        # 데이터 필터링
        filtered_data = df[df['Name'] == name]

        # 시트 이름 생성 (Excel 시트 이름은 최대 31자)
        sheet_name = name[:31]

        # 필터링된 데이터를 별도의 시트에 저장
        filtered_data.to_excel(writer, sheet_name=sheet_name, index=False)

    # Excel 파일 저장 및 닫기
    writer.close()
