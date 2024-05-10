import requests

# URL 설정
url = 'https://192-168-1-242.charlton88.direct.quickconnect.to:5001/oo/r/wnXRY0MgDUqkSC1QiOWdD0wM6JVQzV8Y#tid=3'

# 요청을 보내고 응답을 받음
response = requests.get(url, verify=False)  # SSL 인증을 무시하려면 verify=False를 사용

# 응답 출력
print(response.text)
