import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup

# Chrome 옵션 설정
options = Options()
options.headless = False  # 필요에 따라 True로 설정하여 헤드리스 모드 활성화
options.add_argument("--enable-javascript")  # JavaScript 활성화 (일반적으로 필요 없음)

# WebDriver 설정
driver = webdriver.Chrome(options=options)

# 웹 페이지 로드
driver.get("https://192-168-1-242.charlton88.direct.quickconnect.to:5001/?launchApp=SYNO.SDS.Office.Sheet.Application&launchParam=link%3DwnXRY0MgDUqkSC1QiOWdD0wM6JVQzV8Y#/signin")

# 페이지 로딩을 기다림
time.sleep(3)

# 사용자 이름 입력 필드 찾기
username_input = driver.find_element(By.NAME, "username")
username_input.send_keys("pruchase1")
time.sleep(2)

# 로그인 버튼 클릭
login_button = driver.find_element(By.CLASS_NAME, "login-btn-spinner-wrapper")
login_button.click()
time.sleep(2)

# 비밀번호 입력 필드 찾기
password_input = driver.find_element(By.NAME, "current-password")
password_input.send_keys("Ny8864588")
time.sleep(2)

# 로그인 버튼 클릭
login_button = driver.find_element(By.CLASS_NAME, "login-btn-spinner-wrapper")
login_button.click()
time.sleep(2)

# 페이지의 HTML 정보를 가져오기
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

# 예시: 모든 <div> 태그 출력
divs = soup.find_all('div')
for div in divs:
    print(div.text)


