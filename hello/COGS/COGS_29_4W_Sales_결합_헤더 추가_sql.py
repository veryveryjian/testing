import pandas as pd
import sqlite3

# 데이터베이스 연결 (데이터베이스 파일이 없으면 새로 생성됩니다)
conn = sqlite3.connect('sales_data.db')
cursor = conn.cursor()

# 테이블 생성 (이미 있으면 스킵)
cursor.execute('''
CREATE TABLE IF NOT EXISTS sales_info (
    State TEXT,
    Type TEXT,
    Value REAL
)
''')
conn.commit()

# 파일 경로 설정
file_paths = {
    "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT_SALES.CSV",
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY_SALES.CSV",
    "NJ": "C:\\Users\\charlton\\Desktop\\COGS\\NJ_SALES.CSV",
    "PA": "C:\\Users\\charlton\\Desktop\\COGS\\PA_SALES.CSV",
}

# 결과를 저장할 리스트
results = []

# 파일 읽기 및 데이터 출력
for state, path in file_paths.items():
    df = pd.read_csv(path, encoding='ISO-8859-1', header=1)

    if state == "CT":
        total_inventory_row = df[df.iloc[:, 0] == "Total Inventory"]
        if not total_inventory_row.empty:
            inventory_value = total_inventory_row.iloc[0, 2]
            results.append({'State': state, 'Type': 'Total Inventory', 'Value': inventory_value})
        else:
            results.append({'State': state, 'Type': 'Total Inventory', 'Value': None})
    else:
        if state == "NJ":
            last_row_value = df.iloc[-1]['Unnamed: 10']
        elif state == "NY":
            last_row_value = df.iloc[-1]['Unnamed: 11']
        elif state == "PA":
            last_row_value = df.iloc[-1]['Unnamed: 11']

        results.append({'State': state, 'Type': 'Last Row Specific Column', 'Value': last_row_value})

        last_row_first_col_value = df.iloc[-1, 0]
        results.append({'State': state, 'Type': 'Last Row First Column', 'Value': last_row_first_col_value})

# 결과 데이터프레임 생성
results_df = pd.DataFrame(results)

# 결과를 데이터베이스에 저장
results_df.to_sql('sales_info', conn, if_exists='append', index=False)

# 데이터베이스 연결 종료
conn.close()

# 데이터베이스에 저장된 데이터 확인
conn = sqlite3.connect('sales_data.db')
df = pd.read_sql_query("SELECT * FROM sales_info", conn)
print(df)
conn.close()
