import pandas as pd

# 파일 경로 정의
ny_test_file_path1 = 'C:/Users/charlton/Desktop/popo.xlsx'
#네트워크로부터 해당 데이터를 가져오는 것
#po_____________3_price_list_완성 /// 이것을 우선 실행해서 p_data를 가져올 것
ny_test_file_path2 = 'C:/Users/charlton/Desktop/p.xlsx'

# 첫 번째 파일에서 데이터 읽기 (상위 15행 스킵)
df1 = pd.read_excel(ny_test_file_path1, skiprows=15)

# 두 번째 열이 정수인지 확인 (소수점 없는지)
if not all(df1.iloc[:, 1].apply(lambda x: float(x).is_integer())):
    print("소수점 에러: df1의 두 번째 열에 소수점이 있는 숫자가 있습니다.")
else:
    # 두 번째 파일에서 데이터 읽기
    df2 = pd.read_excel(ny_test_file_path2)

    # df1과 df2를 ITEM 열 기준으로 left join
    result_df = pd.merge(df1, df2[['ITEM', 'Price']], on='ITEM', how='left', suffixes=('_x', '_y'))

    # 계산 열 추가
    result_df['y-x'] = result_df['Price_y'] - result_df['Price_x']
    result_df['y*qty'] = result_df['Price_y'] * df1.iloc[:, 1]  # df1의 두 번째 열을 'Quantity'로 가정
    result_df['y*qty - Amount'] = result_df['y*qty'] - result_df['Amount']

    # total_sum 값 추가
    for col in [ 'y-x', 'y*qty', 'y*qty - Amount']:
        result_df.loc['Total', col] = result_df[col].sum()

    # 결과 출력
    print(result_df)

    # 결과를 엑셀 파일로 저장
    new_file_path = ny_test_file_path1.replace('.xlsx', '_c.xlsx')
    result_df.to_excel(new_file_path, index=False)
    print(f"결과 파일이 저장되었습니다: {new_file_path}")
