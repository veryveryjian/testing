import os
import pandas as pd

# 경로 설정
directory = r'C:\Users\charlton\Desktop\pos2'

# 최종 결과를 저장할 데이터프레임
final_df = pd.DataFrame()

# 해당 디렉토리의 모든 파일을 순회
for filename in os.listdir(directory):
    if filename.endswith(".xls") or filename.endswith(".xlsx"):
        # 파일의 전체 경로를 구성
        filepath = os.path.join(directory, filename)

        try:
            # 엑셀 파일 읽기
            df = pd.read_excel(filepath)

            # 'TOTAL' 문자열을 포함하는 행 찾기
            total_rows = df[df.iloc[:, 0].astype(str).str.contains("total")].copy()  # copy()를 추가하여 명확하게 복사본을 만듬

            # 파일명을 새 열에 추가 (안전하게 .loc 사용)
            total_rows.loc[:, 'Source File'] = filename  # .loc을 사용하여 경고를 방지

            # 최종 데이터프레임에 추가
            final_df = pd.concat([final_df, total_rows], ignore_index=True)

        except Exception as e:
            print(f"파일을 처리하는 중 오류 발생: {filename}")
            print(str(e))

# 최종 데이터프레임을 엑셀 파일로 저장
output_filepath = os.path.join(directory, 'final_output2.xlsx')
final_df.to_excel(output_filepath, index=False)
print(f"최종 결과가 {output_filepath}에 저장되었습니다.")
