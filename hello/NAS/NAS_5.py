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
username_input.send_keys("purchase1")

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
password_input.send_keys("Ny8864588")

# 비밀번호 입력 후 3초 대기
time.sleep(2)

# 로그인 버튼(또는 스피너) 찾기
# 같은 클래스 이름이 사용되었기 때문에 첫 번째 버튼이 다시 클릭될 수 있습니다.
# 다른 버튼을 찾기 위해서는 추가 식별자가 필요할 수 있습니다.
login_button = driver.find_element(By.CLASS_NAME, "login-btn-spinner-wrapper")
login_button.click()

# 버튼 클릭 후 추가로 3초 대기
time.sleep(2)

# 작업 완료 후 드라이버 종료
input()
