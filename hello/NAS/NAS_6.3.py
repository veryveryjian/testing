import os
import time
import urllib.request
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
username_input.send_keys("purchase1")  # 사용자 이름 입력

# 로그인 버튼(또는 스피너) 클릭
login_button = driver.find_element(By.CLASS_NAME, "login-btn-spinner-wrapper")
login_button.click()

# 비밀번호 입력 필드 찾기
time.sleep(2)
password_input = driver.find_element(By.NAME, "current-password")
password_input.send_keys("Ny8864588")  # 비밀번호 입력

# 로그인 버튼(또는 스피너) 클릭
login_button = driver.find_element(By.CLASS_NAME, "login-btn-spinner-wrapper")
login_button.click()

# 페이지의 HTML 정보를 가져오기
time.sleep(2)
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

# 세션 시작
session = requests.Session()

# 로그인 후 페이지로 이동
main_page_url = "https://192-168-1-242.charlton88.direct.quickconnect.to:5001/?launchApp=SYNO.SDS.Drive.Application#file_id=787647127986675835"
response = session.get(main_page_url)

# 파일을 저장할 디렉토리 설정
download_dir = "downloaded_files"
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

# 스타일시트 다운로드
stylesheets = soup.find_all('link', rel='stylesheet')
for link in stylesheets:
    stylesheet_url = link['href']
    filename = os.path.join(download_dir, os.path.basename(stylesheet_url))
    urllib.request.urlretrieve(stylesheet_url, filename)
    print(f"Downloaded: {filename}")

# JavaScript 파일 다운로드
scripts = soup.find_all('script', src=True)
for script in scripts:
    script_url = script['src']
    filename = os.path.join(download_dir, os.path.basename(script_url))
    urllib.request.urlretrieve(script_url, filename)
    print(f"Downloaded: {filename}")

# 작업 완료 후 드라이버 종료
driver.quit()
