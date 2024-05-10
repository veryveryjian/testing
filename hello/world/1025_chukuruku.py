import pandas as pd

# 1. 각 CSV 파일을 pandas 데이터프레임으로 읽기
ny = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku\\ny.CSV", encoding='ISO-8859-1')
nj = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku\\nj.CSV", encoding='ISO-8859-1')
ct = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku\\ct.CSV", encoding='ISO-8859-1')
pa = pd.read_csv("C:\\Users\\charlton\\Desktop\\chukuruku\\pa.CSV", encoding='ISO-8859-1')

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

# 4. 'Name', 'Date', 'Num', 'Credit', 'Memo' 컬럼만 선택하여 'info' 데이터프레임 생성
info = result[['Name', 'Date', 'Num', 'Debit', 'Memo', 'DataFrom']]

# 5. 데이터프레임의 컬럼명을 설정합니다.
columns = ['Name', 'Date', 'Num', 'po#', 'Debit', 'poDate', 'poPrice', 'Memo', 'DataFrom']

# 6. 빈 데이터프레임을 생성합니다.
df = pd.DataFrame(columns=columns)

# 7. 'info' 데이터프레임의 데이터를 'df'에 채워 넣습니다.
df[['Name', 'Date', 'Num', 'Debit', 'Memo', 'DataFrom']] = info

# 8. 엑셀 파일로 저장합니다.
df.to_excel("C:\\Users\\charlton\\Desktop\\chukuruku\\chukuruku_auto.xlsx", index=False)
