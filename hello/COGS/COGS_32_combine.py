import pandas as pd
import os

# 사용자 이름을 여기에 맞게 설정하세요.
username = "charlton"

# 파일 경로 설정
desktop_path = f"C:\\Users\\{username}\\Desktop\\"
output_file_path = os.path.join(desktop_path, "output_combined.xlsx")

file_paths_inventory = {
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY.CSV",
    "NJ": "C:\\Users\\charlton\\Desktop\\COGS\\NJ.CSV",
    "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT.CSV",
    "PA": "C:\\Users\\charlton\\Desktop\\COGS\\PA.CSV"
}

file_paths_sales = {
    "NY": "C:\\Users\\charlton\\Desktop\\COGS\\NY_SALES.CSV",
    "NJ": "C:\\Users\\charlton\\Desktop\\COGS\\NJ_SALES.CSV",
    "CT": "C:\\Users\\charlton\\Desktop\\COGS\\CT_SALES.CSV",
    "PA": "C:\\Users\\charlton\\Desktop\\COGS\\PA_SALES.CSV",
}

# 결과를 저장할 빈 리스트 생성
inventory_info = []

# 첫 번째 데이터 세트 처리
for state, path in file_paths_inventory.items():
    df1 = pd.read_csv(path, encoding='ISO-8859-1', header=2)
    total_inventory_row = df1[df1.iloc[:, 0] == "Total Inventory"]
    df2 = pd.read_csv(path, encoding='ISO-8859-1', header=0)

    inventory_dict = {'State': state}

    if not total_inventory_row.empty:
        row_values = total_inventory_row.iloc[0]
        inventory_dict['AMOUNT'] = row_values['Unnamed: 2']
        inventory_dict['COGS'] = row_values['Unnamed: 5']
    else:
        inventory_dict['AMOUNT'] = None
        inventory_dict['COGS'] = None

    inventory_dict['Header'] = df2.columns.values
    inventory_info.append(inventory_dict)

formatted_inventory_df = pd.DataFrame(inventory_info)

# 두 번째 데이터 세트 처리
results = []

for state, path in file_paths_sales.items():
    df = pd.read_csv(path, encoding='ISO-8859-1', header=1)

    if state == "CT":
        total_inventory_row = df[df.iloc[:, 0] == "Total Inventory"]
        if not total_inventory_row.empty:
            inventory_value = total_inventory_row.iloc[0, 2]
            results.append({'State': state, 'Value': inventory_value, 'First Column': ''})
        else:
            results.append({'State': state, 'Value': None, 'First Column': ''})
    else:
        if state == "NJ":
            last_row_value = df.iloc[-1]['Unnamed: 11']
        elif state == "NY" or state == "PA":
            last_row_value = df.iloc[-1]['Unnamed: 11']

        last_row_first_col_value = df.iloc[-1, 0]
        results.append({'State': state, 'Value': last_row_value, 'First Column': last_row_first_col_value})

results_df = pd.DataFrame(results)
ordered_states = ['NY', 'NJ', 'CT', 'PA']
results_df = results_df.set_index('State').reindex(ordered_states).reset_index()
results_df = results_df[['State', 'Value', 'First Column']]

# 결과를 동일한 엑셀 파일에 다른 시트로 저장
with pd.ExcelWriter(output_file_path) as writer:
    formatted_inventory_df.to_excel(writer, sheet_name='COGS', index=False)
    results_df.to_excel(writer, sheet_name='Sales', index=False)

# 결과 출력
print(f"Formatted Inventory Data saved to {output_file_path} in 'Formatted Inventory' sheet.")
print(f"Sales Data Formatted saved to {output_file_path} in 'Sales Data Formatted' sheet.")

