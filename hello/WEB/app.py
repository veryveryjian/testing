from selenium import webdriver
import chromedriver_autoinstaller

# Chromedriver 자동 설치
chromedriver_autoinstaller.install()

# WebDriver 인스턴스 생성
driver = webdriver.Chrome()

# 이후 웹 페이지를 열거나 다른 동작을 수행하는 코드를 추가
driver.get('https://www.google.com')

# 사용자 입력을 받기 전에 콘솔에 메시지를 표시할 수 있습니다.
print("Press Enter to continue...")
input()

# 모든 작업을 마친 후 드라이버 종료
driver.quit()
