from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options

# geckodriver.exe 파일의 경로를 지정
gecko_path = r'C:\Users\charlton\PycharmProjects\pythonProject\hello\NAS\geckodriver.exe'
service = Service(executable_path=gecko_path)

# 사용자 정의 Firefox 옵션 설정
options = Options()
profile_path = r'C:\Users\charlton\AppData\Roaming\Mozilla\Firefox\Profiles\mwt813y9.default-release'
options.profile = profile_path
options.set_preference("dom.webdriver.enabled", False)
options.set_preference("useAutomationExtension", False)

# Firefox 드라이버 인스턴스 생성
driver = webdriver.Firefox(service=service, options=options)

# 이제 driver 객체를 사용하여 웹 자동화 작업을 수행할 수 있습니다.
