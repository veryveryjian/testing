import sqlite3


# 데이터베이스에 저장된 데이터 확인
conn = sqlite3.connect('sales_data.db')
df = pd.read_sql_query("SELECT * FROM sales_info", conn)
print(df)
conn.close()
