import pdfplumber

# PDF 파일 경로 설정
pdf_path = r"C:\Users\charlton\Desktop\A.pdf"

# 추출할 데이터의 초기 구조 설정
data = {
    'Shipper': [],
    'Consignee': [],
    'MB/L NO': [],
    'Vessel & Voyage No': [],
    'Port of Loading': [],
    'ETA': []
}

# PDF 파일 열기
with pdfplumber.open(pdf_path) as pdf:
    pages = pdf.pages  # 모든 페이지 접근
    for page in pages:
        text = page.extract_text()  # 페이지의 텍스트 추출
        for line in text.split('\n'):  # 각 줄별로 데이터 처리
            if 'SHIPPER:' in line:
                data['Shipper'].append(line.split('SHIPPER:')[1].strip())
            if 'CONSIGNEE:' in line:
                data['Consignee'].append(line.split('CONSIGNEE:')[1].strip())
            if 'MB/L NO' in line:
                data['MB/L NO'].append(line.split('MB/L NO:')[1].strip())
            if 'VESSEL & VOY NO' in line:
                data['Vessel & Voyage No'].append(line.split('VESSEL & VOY NO:')[1].strip())
            if 'PORT OF LOADING' in line:
                data['Port of Loading'].append(line.split('PORT OF LOADING:')[1].strip())
            if 'ETA' in line:
                data['ETA'].append(line.split('ETA:')[1].strip())

# 데이터 출력
for key, value in data.items():
    print(f"{key}: {value}")
