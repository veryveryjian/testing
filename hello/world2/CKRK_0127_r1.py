import pandas as pd

# 생략된 열 없이 모든 열을 출력하도록 설정
pd.set_option('display.max_columns', None)

# 각 지역별 정렬 순서를 지정합니다.
order_ny = ['Chunghua NJ', 'Chunghua CT', 'Charlton PA', 'Miller Cabinetry']
order_nj = ['Charlton Cabinetry', 'Chung Hua Cabinet NJ', 'Chunghua Cabinet CT Inc.', 'Charlton Cabinetry -- PA', 'Miller Cabinetry Inc.']
order_ct = ['Charlton Cabinetry', 'Chung Hua Cabinet NJ', 'Charlton Cabinetry PA', 'Miller Cabinetry Inc']
order_pa = ['CHARLTON CABINETRY NY', 'CHARLTON CABINETRY NJ', 'CHARLTON CABINETRY CT', 'Miller Cabinetry Inc']

# 1. 각 CSV 파일을 pandas 데이터프레임으로 읽기
ny = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\ny.CSV", encoding='ISO-8859-1')
nj = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\nj.CSV", encoding='ISO-8859-1')
ct = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\ct.CSV", encoding='ISO-8859-1')
pa = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\pa.CSV", encoding='ISO-8859-1')

# 2. 'DataFrom' 컬럼 추가 및 데이터프레임 정렬
ny['DataFrom'] = 'NY'
ny['Name'] = pd.Categorical(ny['Name'], categories=order_ny, ordered=True)
ny.sort_values('Name', inplace=True)

nj['DataFrom'] = 'NJ'
nj['Name'] = pd.Categorical(nj['Name'], categories=order_nj, ordered=True)
nj.sort_values('Name', inplace=True)

ct['DataFrom'] = 'CT'
ct['Name'] = pd.Categorical(ct['Name'], categories=order_ct, ordered=True)
ct.sort_values('Name', inplace=True)

pa['DataFrom'] = 'PA'
pa['Name'] = pd.Categorical(pa['Name'], categories=order_pa, ordered=True)
pa.sort_values('Name', inplace=True)

# 3. 모든 데이터프레임을 하나로 합치기
frames = [ny, nj, ct, pa]
result = pd.concat(frames)

# 빈 열 추가하기 (P. O. # 앞에)
result.insert(loc=result.columns.get_loc('P. O. #') if 'P. O. #' in result.columns else 0, column='Empty Column', value='')

# 엑셀 파일로 저장합니다.
with pd.ExcelWriter("C:\\Users\\charlton\\Desktop\\CKRK\\CKRK_auto.xlsx") as writer:
    result.to_excel(writer, index=False)
    for column in result:
        column_length = max(result[column].astype(str).map(len).max(), len(column))
        col_idx = result.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_length)

# 프로세스 종료 메시지
print("Process finished with exit code 0")
