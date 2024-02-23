from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time

import cx_Oracle



options = Options()
options.add_argument('--start-maximized')
options.add_experimental_option("detach", True)
options.add_experimental_option("excludeSwitches", ['enable-logging'])

driver = webdriver.Chrome(options=options)
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": """ Object.defineProperty(navigator, 'webdriver', { get: () => undefined }) """})

# 교보문고 페이지 열기

#교보문고 초등 한페이지 50개 정렬 순으로 50개씩 
#1학년 , 2학년 국어, 통합교과(바슬즐), 수학, 예체능, 전과목
#39010101 초등참고서 > 학년별 개념 / 문제 > 초등1학년 > 국어
#39010102 초등참고서 > 학년별 개념 / 문제 > 초등1학년 > 통합교과(바슬즐)
#39010107 초등참고서 > 학년별 개념 / 문제 > 초등1학년 > 수학
#39010115 초등참고서 > 학년별 개념 / 문제 > 초등1학년 > 예체능
#39010117 초등참고서 > 학년별 개념 / 문제 > 초등1학년 > 전과목

#39010301 초등참고서 > 학년별 개념 / 문제 > 초등2학년 > 국어
#39010302 초등참고서 > 학년별 개념 / 문제 > 초등2학년 > 통합교과(바슬즐)
#39010307 초등참고서 > 학년별 개념 / 문제 > 초등2학년 > 수학
#39010315 초등참고서 > 학년별 개념 / 문제 > 초등2학년 > 예체능
#39010317 초등참고서 > 학년별 개념 / 문제 > 초등2학년 > 전과목

#3학년, 4학년, 5학년, 6학년 국어, 수학, 사회, 과학, 영어, 예체능, 전과목
#39010501 초등참고서 > 학년별 개념 / 문제 > 초등3학년 > 국어
#39010507 초등참고서 > 학년별 개념 / 문제 > 초등3학년 > 수학
#39010509 초등참고서 > 학년별 개념 / 문제 > 초등3학년 > 사회
#39010511 초등참고서 > 학년별 개념 / 문제 > 초등3학년 > 과학
#39010515 초등참고서 > 학년별 개념 / 문제 > 초등3학년 > 영어
#39010521 초등참고서 > 학년별 개념 / 문제 > 초등3학년 > 예체능
#39010523 초등참고서 > 학년별 개념 / 문제 > 초등3학년 > 전과목

#39010701 초등참고서 > 학년별 개념 / 문제 > 초등4학년 > 국어
#39010707 초등참고서 > 학년별 개념 / 문제 > 초등4학년 > 수학
#39010709 초등참고서 > 학년별 개념 / 문제 > 초등4학년 > 사회
#39010711 초등참고서 > 학년별 개념 / 문제 > 초등4학년 > 과학
#39010715 초등참고서 > 학년별 개념 / 문제 > 초등4학년 > 영어
#39010721 초등참고서 > 학년별 개념 / 문제 > 초등4학년 > 예체능
#39010723 초등참고서 > 학년별 개념 / 문제 > 초등4학년 > 전과목

#39010901 초등참고서 > 학년별 개념 / 문제 > 초등5학년 > 국어
#39010907 초등참고서 > 학년별 개념 / 문제 > 초등5학년 > 수학
#39010909 초등참고서 > 학년별 개념 / 문제 > 초등5학년 > 사회
#39010911 초등참고서 > 학년별 개념 / 문제 > 초등5학년 > 과학
#39010915 초등참고서 > 학년별 개념 / 문제 > 초등5학년 > 영어
#39010921 초등참고서 > 학년별 개념 / 문제 > 초등5학년 > 예체능
#39010923 초등참고서 > 학년별 개념 / 문제 > 초등5학년 > 전과목

#39011101 초등참고서 > 학년별 개념 / 문제 > 초등6학년 > 국어
#39011107 초등참고서 > 학년별 개념 / 문제 > 초등6학년 > 수학
#39011109 초등참고서 > 학년별 개념 / 문제 > 초등6학년 > 사회
#39011111 초등참고서 > 학년별 개념 / 문제 > 초등6학년 > 과학
#39011115 초등참고서 > 학년별 개념 / 문제 > 초등6학년 > 영어
#39011121 초등참고서 > 학년별 개념 / 문제 > 초등6학년 > 예체능
#39011123 초등참고서 > 학년별 개념 / 문제 > 초등6학년 > 전과목


url = f'https://product.kyobobook.co.kr/category/KOR/39010101#?page=1&type=all&sort=sel'
driver.get(url)
time.sleep(2)

# 스크롤 다운 (페이지의 끝까지 스크롤)
last_height = driver.execute_script("return document.body.scrollHeight")
while True:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(3)  # 스크롤 동안 대기
    new_height = driver.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height


# 스크래핑
book_id = driver.find_elements(By.CLASS_NAME,'prod_item')
book_title = driver.find_elements(By.CLASS_NAME, 'prod_name')
book_author = driver.find_elements(By.CLASS_NAME, 'prod_author')
book_discounted_rate = driver.find_elements(By.CSS_SELECTOR, '.prod_price .percent')

book_discounted_price = driver.find_elements(By.CLASS_NAME, 'price')
book_normal_price = driver.find_elements(By.CSS_SELECTOR,'.price_normal .val')
book_image = driver.find_elements(By.CSS_SELECTOR, '.prod_item .img_box')
book_introduction = driver.find_elements(By.CLASS_NAME, 'prod_introduction')


for id, title, author, discounted_rate, discounted_price, normal_price, image, introduction in zip(book_id, book_title, book_author, book_discounted_rate, book_discounted_price, book_normal_price, book_image, book_introduction) :
    id_text = id.get_attribute('data-id')
    title_text = title.text
    #prod_author '이미례', '리틀씨앤톡', '2024.02.15' 데이터구조 
    author_text = author.text.split(" · ")[0]
    publisher_text = author.text.split(" · ")[1]
    publish_date_text =author.text.split(" · ")[2]
    discounted_rate_text = discounted_rate.text
    price_text = discounted_price.text
    normal_price_text = normal_price.text
    introduction_text = introduction.text
    image_url = image.find_element(By.TAG_NAME, 'img').get_attribute('src') if book_image else "이미지 URL 없음"
    print("상품ID : ",id_text)
    print("제목 : ",title_text)
    print("저자 : ",author_text)
    print("출판사 : ",publisher_text)
    print("출판일 : ",publish_date_text)
    print("할인율 : ",discounted_rate_text)
    print("할인가격 : ",price_text)
    print("정가 ",normal_price_text)

    print("소개 : ",introduction_text)
    print("이미지 : ",image_url)
    print("\n")


time.sleep(5)

# 브라우저 닫기
driver.quit()

#DB 접속

con = cx_Oracle.connect("HWAYEON", "HWAYEON", "192.168.0.122:1521/xe", encoding="UTF-8") #오라클 연결
print( con ) #연결확인

cursor = con.cursor() #CRUD명령 실행을 위한 커서 객체를 얻는다.
#쿼리문
sql = """select * from boards"""
cursor.execute(sql)
x = cursor.fetchall()

print("=====>", x)
con.commit() #커밋을 통한 트랜잭션 종료

cursor.close()
con.close()