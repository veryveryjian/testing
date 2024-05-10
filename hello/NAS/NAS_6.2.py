import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import requests

# 웹드라이버 설정
driver = webdriver.Chrome()

# 로그인 페이지로 이동
driver.get("https://192-168-1-242.charlton88.direct.quickconnect.to:5001/?launchApp=SYNO.SDS.Drive.Application#file_id=787647127986675835")

# 페이지 로딩을 기다림
time.sleep(3)

# 사용자 이름 입력 필드 찾기
username_input = driver.find_element(By.NAME, "username")

# 'abc' 입력
username_input.send_keys("aaa")

# 사용자 이름 입력 후 3초 대기
time.sleep(2)

# 로그인 버튼(또는 스피너) 찾기
login_button = driver.find_element(By.CLASS_NAME, "login-btn-spinner-wrapper")
login_button.click()

# 버튼 클릭 후 추가로 3초 대기
time.sleep(2)

# 비밀번호 입력 필드 찾기
password_input = driver.find_element(By.NAME, "current-password")

# 'abc' 입력
password_input.send_keys("aaa")

# 비밀번호 입력 후 3초 대기
time.sleep(2)

# 로그인 버튼(또는 스피너) 찾기
login_button = driver.find_element(By.CLASS_NAME, "login-btn-spinner-wrapper")
login_button.click()

# 버튼 클릭 후 추가로 3초 대기
time.sleep(2)

# 페이지의 HTML 정보를 가져오기
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

# 세션 시작
session = requests.Session()

# 로그인 후 페이지로 이동
main_page_url = "https://192-168-1-242.charlton88.direct.quickconnect.to:5001/?launchApp=SYNO.SDS.Drive.Application#file_id=787647127986675835"
response = session.get(main_page_url)

# 페이지의 HTML 출력
print(response.text)

# 작업 완료 후 드라이버 종료하지 않음
# driver.quit()
