import pandas as pd

# 1단계: Hello_NY.csv 처리
file_path = 'C:/Users/charlton/Desktop/Hello_NY.csv'
df = pd.read_csv(file_path)
split_data = df['Unnamed: 0'].str.split('-', expand=True).fillna('')
df['Color_Code'] = split_data[0]
df['Model_Code'] = split_data[1]

# 2단계: CBM_test.xlsx와 결합
cbm_path = 'C:/Users/charlton/Desktop/CBM_test.xlsx'
cbm_df = pd.read_excel(cbm_path)
NY_cbm = pd.merge(df, cbm_df, on='Model_Code', how='left')

# 3단계: MS_avg.xlsx와 결합
avg_path = 'C:/Users/charlton/Desktop/MS_avg.xlsx'
avg_df = pd.read_excel(avg_path)
NY_cbm_avg = pd.merge(NY_cbm, avg_df, left_on='Unnamed: 0', right_on='ITEM', how='left')

# 4단계: MS-room.xlsx와 결합
room_path = 'C:/Users/charlton/Desktop/MS-room.xlsx'
room_df = pd.read_excel(room_path)
final_df = pd.merge(NY_cbm_avg, room_df, left_on='Unnamed: 0', right_on='ITEM', how='left')

# MS에 해당하는 Color_Code 값만 필터링
ms_filtered_df = final_df[final_df['Color_Code'] == 'MS']

# 선택된 열만 포함하여 최종 DataFrame 조정
selected_columns = ms_filtered_df[['Unnamed: 0', 'On Hand', 'NY']]

# 결과를 ny_sim2.xlsx로 저장
final_path = 'C:/Users/charlton/Desktop/ny_sim2.xlsx'
selected_columns.to_excel(final_path, index=False)




