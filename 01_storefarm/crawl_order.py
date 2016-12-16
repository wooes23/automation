# -*- coding:utf8 -*-
"""
1.crawl_data(selenium 활용)
    1. storefarm에 접속하고,
    2. 신규 주문내역이 있는지 확인한 뒤
    3. 있으면 전체 주문내역 다운로드(엑셀파일)
        a. 다운로드 경로 지정
    4. 다운로드 완료하고 주문내역 선택한 뒤 발주 확인
"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys

userID = raw_input("ID : ")
userPW = raw_input("password : ")

# chromedriver를 이용하여 storefarm Login 창에 접속합니다.
path = './chromedriver'
driver = webdriver.Chrome(path)
driver.get('https://sell.storefarm.naver.com/l/login/')


# Login
# ID와 PW 영역을 찾아줍니다
el_name = driver.find_element_by_name('username')
el_pw = driver.find_element_by_name('password')

# userID 입력
el_name.clear()
el_name.send_keys(userID)

# password 입력
el_pw.clear()
el_pw.send_keys(userPW)
el_pw.send_keys(Keys.RETURN)


#신규 주문내역이 있는지 확인
el_new = driver.find_element_by_class_name('_new_order')
text_new_order = int(el_new.text)

# 신규 주문이 있을 경우
if text_new_order > 0 :
    el_new.click()

    # 전체 주문내역 다운로드
    el_order_down = driver.find_element_by_class_name('_excelDownloadBtn')
    el_order_down.click()

    # 다운로드 경로는 바탕화면


    # 주문이 몇개인지 확인
    print 'new order :'+ str(text_new_order)

else :
    print 'new order :'+ str(text_new_order)



# 다운로드 완료 후 주문내역을 선택한 뒤 발주 확인
