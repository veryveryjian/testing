from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ChromeDriver가 시스템의 PATH에 포함되어 있다고 가정하고, 별도의 경로 지정 없이 인스턴스 생성
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 10)

try:
    # Google 로그인 페이지로 이동
    driver.get("https://accounts.google.com/signin")

    # 이메일 입력
    email_input = wait.until(EC.element_to_be_clickable((By.ID, "identifierId")))
    email_input.send_keys("seungwoousa@gmail.com")  # 실제 사용할 이메일 주소로 변경

    # '다음' 버튼 클릭
    next_button = driver.find_element(By.ID, "identifierNext")
    next_button.click()

    # 패스워드 입력을 위해 필드가 활성화될 때까지 대기
    password_input = wait.until(EC.element_to_be_clickable((By.NAME, "password")))
    password_input.send_keys("your_password")  # 실제 사용할 패스워드로 변경

    # '다음' 버튼을 다시 클릭하여 로그인
    next_button = driver.find_element(By.ID, "passwordNext")
    next_button.click()

    # 로그인이 성공했는지 확인하는 로직을 추가할 수 있습니다.

finally:
    # 작업 완료 후 브라우저 종료
    driver.quit()
