import pandas as pd

# Excel 파일 경로
excel_file = r'C:\Users\charlton\Desktop\test1.xlsx'

# Excel 파일 전체 로드 (INVOICE 시트, 전체 데이터)
df = pd.read_excel(excel_file, sheet_name='INVOICE')

# 열 이름 확인
print("열 이름:", df.columns.tolist())

# 'Item' 열 존재 여부 확인
if 'Item' in df.columns:
    # 'Item' 열을 기준으로 헤더 위치 식별
    header_indices = df.index[df['Item'] == 'Item'].tolist()

    # 데이터 구간 분리 및 변수 할당
    sections = {}  # 각 섹션을 저장할 딕셔너리
    for i, start in enumerate(header_indices):
        end = header_indices[i + 1] if i + 1 < len(header_indices) else None
        section_df = df[start:end].reset_index(drop=True).iloc[1:]  # 헤더 행 제외
        sections[f'section_{i + 1}'] = section_df

    # 각 섹션 출력
    for name, section_df in sections.items():
        print(f"{name}:")
        print(section_df)
        print("\n")

else:
    print("'Item' 열이 데이터프레임에 존재하지 않습니다.")


print(df)
