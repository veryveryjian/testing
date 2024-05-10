import pandas as pd

# 파일 경로 정의
ny_test_file_path1 = 'C:/Users/charlton/Desktop/po/popo.xlsx'
ny_test_file_path2 = 'C:/Users/charlton/Desktop/po/p.xlsx'

# 첫 번째 파일에서 데이터 읽기 (상위 15행 스킵)
df1 = pd.read_excel(ny_test_file_path1, skiprows=15)

# 두 번째 열이 정수인지 확인 (소수점 없는지)
if not all(df1.iloc[:, 1].apply(lambda x: float(x).is_integer())):
    print("소수점 에러: df1의 두 번째 열에 소수점이 있는 숫자가 있습니다.")
    print("소수점 에러: df1의 두 번째 열에 소수점이 있는 숫자가 있습니다.")
    print("소수점 에러: df1의 두 번째 열에 소수점이 있는 숫자가 있습니다.")
    print("소수점 에러: df1의 두 번째 열에 소수점이 있는 숫자가 있습니다.")
    print("소수점 에러: df1의 두 번째 열에 소수점이 있는 숫자가 있습니다.")
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

    # 'df1 Quantity' 열 추가 (df1의 두 번째 열 데이터)
    result_df['df1 Quantity'] = df1.iloc[:, 1]

    # 'Total' 문자열을 포함하지 않는 행을 필터링하여 'df1 Quantity' 열의 합계 계산
    quantity_sum = result_df[~result_df['ITEM'].str.contains('total', case=False, na=False)]['df1 Quantity'].sum()

    # 각 계산된 열에 대해 합계를 'Total' 행에 추가
    result_df.loc['Total', 'df1 Quantity'] = quantity_sum
    for col in [ 'y-x', 'y*qty', 'y*qty - Amount']:
        result_df.loc['Total', col] = result_df[col].sum(skipna=True)


    # 결과 출력
    print(result_df)

    # 결과를 엑셀 파일로 저장
    new_file_path = ny_test_file_path1.replace('.xlsx', '_c.xlsx')
    result_df.to_excel(new_file_path, index=False)
    print(f"결과 파일이 저장되었습니다: {new_file_path}")
