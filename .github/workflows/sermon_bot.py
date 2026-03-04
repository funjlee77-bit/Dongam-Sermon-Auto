import requests
from bs4 import BeautifulSoup
import os

# 1. 설교 페이지 주소
SERMON_PAGE = "https://www.dongamchurch.kr/EZ/rb/list.asp?BoardModule=Media&tbcode=tv01"

def get_latest_sermon():
    res = requests.get(SERMON_PAGE)
    soup = BeautifulSoup(res.text, "html.parser")

    link = soup.find("a", href=True)
    if link:
        return "https://www.dongamchurch.kr" + link["href"]
    return None

def extract_vimeo_url(page_url):
    res = requests.get(page_url)
    soup = BeautifulSoup(res.text, "html.parser")

    iframe = soup.find("iframe")
    if iframe:
        return iframe["src"]
    return None

def main():
    print("설교 페이지 확인 중...")
    latest_page = get_latest_sermon()
    print("최신 설교 페이지:", latest_page)

    if latest_page:
        vimeo_url = extract_vimeo_url(latest_page)
        print("Vimeo 주소:", vimeo_url)
    else:
        print("설교를 찾지 못했습니다.")

if __name__ == "__main__":
    main()
