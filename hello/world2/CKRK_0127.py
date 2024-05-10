import pandas as pd

# 생략된 열 없이 모든 열을 출력하도록 설정
pd.set_option('display.max_columns', None)

# 1. 각 CSV 파일을 pandas 데이터프레임으로 읽기
ny = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\ny.CSV", encoding='ISO-8859-1')
nj = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\nj.CSV", encoding='ISO-8859-1')
ct = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\ct.CSV", encoding='ISO-8859-1')
pa = pd.read_csv("C:\\Users\\charlton\\Desktop\\CKRK\\pa.CSV", encoding='ISO-8859-1')

# 1.5 각 데이터프레임을 'Name' 컬럼으로 오름차순 정렬
ny = ny.sort_values(by='Name')
nj = nj.sort_values(by='Name')
ct = ct.sort_values(by='Name')
pa = pa.sort_values(by='Name')

# 2. 'DataFrom' 컬럼 추가
ny['DataFrom'] = 'NY'
nj['DataFrom'] = 'NJ'
ct['DataFrom'] = 'CT'
pa['DataFrom'] = 'PA'

# 3. 모든 데이터프레임을 하나로 합치기
frames = [ny, nj, ct, pa]
result = pd.concat(frames)

# 4. 'Name', 'Date', 'Num', 'P. O. #', 'Debit', 'Memo', 'DataFrom' 컬럼만 선택하여 'info' 데이터프레임 생성
info = result[['Name', 'Date', 'Num', 'P. O. #', 'Debit', 'Memo', 'DataFrom']]

# 빈 열 추가하기 (P. O. # 앞에)
df = info.copy()
df.insert(loc=df.columns.get_loc('P. O. #'), column='Blank', value='')

# 8. 엑셀 파일로 저장합니다.
with pd.ExcelWriter("C:\\Users\\charlton\\Desktop\\CKRK\\CKRK_auto.xlsx") as writer:
    df.to_excel(writer, index=False)
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_length)

# 프로세스 종료 메시지
print("Process finished with exit code 0")
