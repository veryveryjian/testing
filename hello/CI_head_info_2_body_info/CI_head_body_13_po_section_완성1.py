import pandas as pd
import os

# 출력 제한 해제
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# 읽어올 파일들이 있는 디렉터리 설정
directory = r'C:\Users\charlton\Desktop\CI'

# 결과 데이터를 저장할 리스트 초기화
results = []

# 디렉터리 내의 모든 Excel 파일 이름을 리스트로 저장
excel_files = [filename for filename in os.listdir(directory) if filename.endswith('.xlsx') and not filename.startswith('~$')]

# 각 Excel 파일을 순회하며 처리
for file in excel_files:
    file_path = os.path.join(directory, file)
    try:
        print(f"\nProcessing File: {file}")
        df = pd.read_excel(file_path, engine='openpyxl')  # 파일에서 데이터 읽기와 엔진 지정

        # 'PO#'로 시작하는 각 주문 섹션의 시작 인덱스 찾기
        po_indices = df.index[df.iloc[:, 0].str.contains("PO#", na=False)].tolist()

        # 섹션 이름 모으기
        section_names = []
        for i, start_index in enumerate(po_indices):
            end_index = po_indices[i + 1] if i + 1 < len(po_indices) else len(df)
            section_name = str(df.iloc[start_index, 0])
            section_names.append(section_name)

        # 결과 리스트에 파일 및 섹션 별 결과 추가
        results.append({
            'File': file,
            'Section Count': len(po_indices),
            'Section Details': ", ".join(section_names)
        })
    except Exception as e:
        print(f"Failed to process {file}: {e}")

# 결과 데이터프레임 생성
results_df = pd.DataFrame(results)

# 결과 출력
print(results_df)

# 결과 데이터프레임을 엑셀 파일로 저장
output_path = r'C:\Users\charlton\Desktop\Total_Results__.xlsx'
results_df.to_excel(output_path, index=False)

# 저장된 파일 경로 출력
print(f"Results have been saved to {output_path}")
