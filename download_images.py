import os
import requests
from bs4 import BeautifulSoup

# 이미지 저장 폴더 생성
os.makedirs('images', exist_ok=True)

# 크롤링할 사이트 주소
url = "https://sites.google.com/view/circlemain/home"

# 사이트 HTML 가져오기
response = requests.get(url)
soup = BeautifulSoup(response.text, 'html.parser')

# 이미지 태그에서 src 추출
img_tags = soup.find_all('img')
img_urls = []
for img in img_tags:
    src = img.get('src')
    if src and src.startswith('http'):
        img_urls.append(src)

# 이미지 다운로드
def download_image(img_url):
    filename = img_url.split('/')[-1].split('?')[0]
    filepath = os.path.join('images', filename)
    try:
        img_data = requests.get(img_url).content
        with open(filepath, 'wb') as handler:
            handler.write(img_data)
        print(f"저장 완료: {filename}")
    except Exception as e:
        print(f"실패: {img_url} ({e})")

for img_url in img_urls:
    download_image(img_url) 