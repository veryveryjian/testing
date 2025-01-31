import pandas as pd
import glob
import os

# 출력 제한 설정 변경
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# 파일 경로 설정
ci_folder_path = r'C:\Users\charlton\Desktop\CI\*.xlsx'
cbm_file_path = r'C:\Users\charlton\Desktop\CBM.xlsx'
CBM_df = pd.read_excel(cbm_file_path)  # CBM 파일 미리 읽기

# 모든 요약 데이터를 저장할 DataFrame 초기화
all_summary_data = pd.DataFrame()

for excel_file in glob.glob(ci_folder_path):
    try:
        df = pd.read_excel(excel_file, sheet_name='INVOICE', skiprows=15, usecols='B:G')

        if 'Item' not in df.columns:
            print(f"Error: 'Item' column not found in {excel_file}. Check the file format.")
            continue  # 'Item' 열이 없으면 다음 파일로 넘어감

        Item_data = df[df['Item'].notna()]
        index_null_check = Item_data.dropna(subset=Item_data.columns[:6]).copy()

        split_data = index_null_check['Item'].str.split('-', n=1, expand=True)
        index_null_check['Color_code'] = split_data[0]
        index_null_check['spec_code'] = split_data[1]

        merged_df = pd.merge(index_null_check, CBM_df, on='spec_code', how='left')
        merged_df['cbmXqty'] = merged_df['CBM'] * merged_df['Pcs']

        new_file_path = os.path.join(r'C:\Users\charlton\Desktop',
                                     os.path.basename(excel_file).replace('.xlsx', '_c.xlsx'))
        with pd.ExcelWriter(new_file_path, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='Joined Data')
            summary_df = merged_df.groupby('Color_code')['cbmXqty'].sum().reset_index()
            summary_df.to_excel(writer, index=False, sheet_name='Summary by Color')

        print(f"Processed and saved: {new_file_path}")

        # 각 파일의 요약 데이터 읽기 및 'dataFrom' 컬럼 추가
        summary_df['dataFrom'] = os.path.basename(new_file_path)
        all_summary_data = pd.concat([all_summary_data, summary_df], ignore_index=True)

    except ValueError as e:
        print(f"Error processing {excel_file}: {e}")

# 전체 요약 데이터를 엑셀 파일로 저장
total_cbm_list_path = r'C:\Users\charlton\Desktop\total_cbm_list.xlsx'
all_summary_data.to_excel(total_cbm_list_path, index=False)
print(f"All summary data saved to {total_cbm_list_path}")
