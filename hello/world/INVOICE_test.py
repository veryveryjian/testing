import pandas as pd

# 각 주별 파일 읽기 및 데이터프레임에 할당
ny_invoice = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku_data 2\\ny_inv.CSV", encoding='ISO-8859-1')
nj_invoice = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku_data 2\\nj_inv.CSV", encoding='ISO-8859-1')
ct_invoice = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku_data 2\\ct_inv.CSV", encoding='ISO-8859-1')
pa_invoice = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku_data 2\\pa_inv.CSV", encoding='ISO-8859-1')


# 각 데이터프레임에 대해 'Num' 컬럼으로 데이터 오름차순 정렬
ny_invoice.sort_values(by='Num', inplace=True)
nj_invoice.sort_values(by='Num', inplace=True)
ct_invoice.sort_values(by='Num', inplace=True)
pa_invoice.sort_values(by='Num', inplace=True)

# NY, NJ, CT, PA의 필터링할 이름들
filter_names = {
    'NY': ['Chunghua NJ', 'Chunghua CT', 'Charlton PA', 'Miller Cabinetry'],
    'NJ': ['Charlton Cabinetry', 'Chunghua Cabinet CT Inc.', 'Charlton Cabinetry -- PA', 'Miller Cabinetry Inc.', 'Chung Hua Cabinet NJ'],
    'CT': ['Charlton Cabinetry', 'Chung Hua Cabinet NJ', 'Charlton Cabinetry PA', 'Miller Cabinetry Inc'],
    'PA': ['CHARLTON CABINETRY NY', 'CHARLTON CABINETRY NJ', 'CHALRTON CABINETRY CT', 'Miller Cabinetry Inc']
}

# 각 주별 데이터 필터링 및 별도의 파일로 저장
for state, names in filter_names.items():
    state_df = locals()[f"{state.lower()}_invoice"]  # 해당 주의 데이터프레임 선택
    state_writer = pd.ExcelWriter(f'/Users\\charlton\\Desktop\\{state.lower()}_invoice_data.xlsx', engine='xlsxwriter')  # 해당 주의 파일 작성자 생성

    for name in names:
        filtered_data = state_df[state_df['Name'].str.contains(name, na=False)]
        sheet_name = f"{name[:31]}"
        filtered_data.to_excel(state_writer, sheet_name=sheet_name, index=False)

    # 각 주별 파일 저장 및 닫기
    state_writer.close()
