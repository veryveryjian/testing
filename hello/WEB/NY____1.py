from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

options = Options()


user_data = r"C:\Users\charlton\AppData\Local\Google\Chrome\User Data\Profile 10"
options.add_argument(f"user-data-dir={user_data}")

options.add_experimental_option("detach", True) # 화면이 꺼지지 않고 유지

options. add_argument("--start-maximized")

service = Service(ChromeDriverManager().install())

driver = webdriver.Chrome(service=service, options=options)

driver.get("https://google.com")