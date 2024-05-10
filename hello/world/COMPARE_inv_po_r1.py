import pandas as pd

# 파일 읽기
inv_path = "C:\\Users\\charlton\\Desktop\\TEST\\INV.CSV"
inv = pd.read_csv(inv_path, encoding='ISO-8859-1')

# 'Num' 열의 값이 53273.0이고, 'Account' 열의 값이 'Sales'인 행 필터링
inv_df = inv[(inv['Num'] == 53273.0) & (inv['Account'] == 'Sales')]

# 필터링된 행 중 'Credit' 및 'Qty' 열의 합계 계산
sum_credit = inv_df['Credit'].sum()
sum_qty = inv_df['Qty'].sum()

# PO 데이터 로드 및 필터링
po_path = "C:\\Users\\charlton\\Desktop\\TEST\\PO.CSV"
po = pd.read_csv(po_path, encoding='ISO-8859-1')
po_df = po[po['Num'] == 53273.0]

# Excel 파일로 저장
output_path = "C:\\Users\\charlton\\Desktop\\Output.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # 'INVOICE' 시트로 저장
    inv_df.to_excel(writer, sheet_name='INVOICE', index=False)

    # 'PO' 시트로 저장
    po_df.to_excel(writer, sheet_name='PO', index=False)

# 'Credit'과 'Qty'의 합계를 새로운 시트에 저장
with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    summary_df = pd.DataFrame({
        'Sum of Credit': [sum_credit],
        'Sum of Qty': [sum_qty]
    })
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
