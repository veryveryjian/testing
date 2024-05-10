import time
from selenium import webdriver
from selenium.webdriver.common.by import By

# 웹드라이버 설정
driver = webdriver.Chrome()

# 로그인 페이지로 이동
driver.get("https://192-168-1-242.charlton88.direct.quickconnect.to:5001/?launchApp=SYNO.SDS.Office.Sheet.Application&launchParam=link%3DwnXRY0MgDUqkSC1QiOWdD0wM6JVQzV8Y#/signin")

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

###################################################################

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

# 페이지의 HTML 정보를 출력
print(driver.page_source)

