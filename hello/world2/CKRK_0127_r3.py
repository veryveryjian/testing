import pandas as pd

# 생략된 열 없이 모든 열을 출력하도록 설정
pd.set_option('display.max_columns', None)

# 각 지역별 정렬 순서를 지정합니다.
order_ny = ['Chunghua NJ', 'Chunghua CT', 'Charlton PA', 'Miller Cabinetry']
order_nj = ['Charlton Cabinetry', 'Chung Hua Cabinet NJ', 'Chunghua Cabinet CT Inc.', 'Charlton Cabinetry -- PA',
            'Miller Cabinetry Inc.']
order_ct = ['Charlton Cabinetry', 'Chung Hua Cabinet NJ', 'Charlton Cabinetry PA', 'Miller Cabinetry Inc']
order_pa = ['CHARLTON CABINETRY NY', 'CHARLTON CABINETRY NJ', 'CHARLTON CABINETRY CT', 'Miller Cabinetry Inc']

# 1. 각 CSV 파일을 pandas 데이터프레임으로 읽기
ny = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\ny.CSV", encoding='ISO-8859-1')
nj = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\nj.CSV", encoding='ISO-8859-1')
ct = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\ct.CSV", encoding='ISO-8859-1')
pa = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\pa.CSV", encoding='ISO-8859-1')


# 'DataFrom' 컬럼 추가 및 데이터프레임 정렬, 그리고 각 지역별 데이터프레임을 Categorical 타입으로 변환
def sort_and_categorize(df, order, state):
    df['DataFrom'] = state
    df['Name'] = pd.Categorical(df['Name'], categories=order, ordered=True)
    return df.sort_values('Name')


ny = sort_and_categorize(ny, order_ny, 'NY')
nj = sort_and_categorize(nj, order_nj, 'NJ')
ct = sort_and_categorize(ct, order_ct, 'CT')
pa = sort_and_categorize(pa, order_pa, 'PA')

# Define the desired column order
desired_column_order = ['Date', 'Num', 'Name', 'P. O. #', 'Debit', 'DataFrom',]

# Reorder the columns of each DataFrame according to the desired column order
ny = ny.reindex(columns=desired_column_order)
nj = nj.reindex(columns=desired_column_order)
ct = ct.reindex(columns=desired_column_order)
pa = pa.reindex(columns=desired_column_order)

# 엑셀 파일로 저장합니다.
excel_path = "C:\\Users\\charlton\\Desktop\\CKRK\\CKRK_test.xlsx"
with pd.ExcelWriter(excel_path) as writer:
    ny.to_excel(writer, sheet_name='NY', index=False)
    nj.to_excel(writer, sheet_name='NJ', index=False)
    ct.to_excel(writer, sheet_name='CT', index=False)
    pa.to_excel(writer, sheet_name='PA', index=False)

    # 각 열의 너비를 시트별로 조정
    for state_df, state in zip([ny, nj, ct, pa], ['NY', 'NJ', 'CT', 'PA']):
        for column in desired_column_order:  # Ensure we only iterate over the desired columns
            if column in state_df:
                column_length = max(state_df[column].astype(str).map(len).max(), len(column))
                col_idx = state_df.columns.get_loc(column)
                writer.sheets[state].set_column(col_idx, col_idx, column_length)

# 프로세스 종료 메시지
print("Process finished with exit code 0")
