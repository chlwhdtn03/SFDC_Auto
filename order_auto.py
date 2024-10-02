#!/usr/bin/env python
# coding: utf-8

# In[1]:


import selenium
from selenium import webdriver
from selenium.webdriver import ActionChains

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

import xlwings as xw
import pandas as pd
from xlwings.constants import DeleteShiftDirection

from selenium.common.exceptions import NoSuchElementException

from win11toast import toast


# In[2]:


print("[정보] Office 정품 인증이 받아진 상태로 실행해야 문제 없이 실행될 수 있습니다.")
print("")
input("[정보] Enter를 눌러 order.xlsx를 불러옵니다.")


# In[47]:


book = xw.Book("order.xlsx")
sheet = book.sheets[0]
sheet.range('3:3').api.Delete(DeleteShiftDirection.xlShiftUp) 
target = sheet.used_range.options(pd.DataFrame, index=False).value
book.close()
urlist = list(filter(None,target['URL']))
print("[정보] order.xlsx를 불러왔습니다.")

def refresh_excel():
    global book, sheet, target, urlist
    book = xw.Book("order.xlsx")
    sheet = book.sheets[0]
    sheet.range('3:3').api.Delete(DeleteShiftDirection.xlShiftUp) 
    target = sheet.used_range.options(pd.DataFrame, index=False).value
    book.close()
    urlist = list(filter(None,target['URL']))


# In[14]:


service = webdriver.ChromeService(executable_path='chromedriver/chromedriver.exe')

driver = webdriver.Chrome(service=service)


# In[15]:


url = urlist[0] # 주문 조회페이지로 바로 접속
driver.get(url)
input("[정보] 로그인이 되었으면 Enter를 눌러주세요")


# In[16]:


driver.get(url)
input("[정보] 공지창이 열려있다면 7일 동안 다시보지 않기 버튼을 꼭 눌러주세요! (Enter키를 눌러 계속)")


# In[49]:


# 이미 주문번호가 발행된 건지 확인
wait = WebDriverWait(driver, 60)
error = 0
SET_AMOUNT = 5
current_url = ''
current_pos = 0
selected = ''
def process():
    global error, SET_AMOUNT, current_url, current_pos, selected
    if(error > 0):
        print("[정보] 작업을 다시 시작합니다..")
        refresh_excel()
    error = 0
    selected = ''
    current_pos = 0
    
    for url in urlist:
        if error:
            chkERR()
            return
        current_pos = 0 # url 이동시 마다 초기화 필수
        current_url = url
        driver.get(url)
        
        while True:
            try: 
                selected = driver.find_element(By.ID, 'j_id0:mainFrm:multipleOrder:outMultipleOrderList:%d:j_id86' % current_pos).text
            except NoSuchElementException:
                if current_pos > 0:
                    break
                elif current_pos == 0:
                    error = 3
                    break
                
            if selected == 'ERROR':
                error = 2
                break
            for i in range(SET_AMOUNT):
                if selected != driver.find_element(By.ID, 'j_id0:mainFrm:multipleOrder:outMultipleOrderList:%d:j_id86' % (current_pos+i)).text:
                    error = 1
                    break
            if selected != '' and not error: # 계속해서 탐색 & 방금 직전에 주문저장해서 새로 발급된 주문번호에 문제가 없으면 POS 이동
                current_pos += SET_AMOUNT
            elif error: # 불일치 에러 발생
                chkERR()
                return
            elif selected == '': 
                # 주문번호가 모두 빈칸인 곳까지 내려온 상태
                # 체크박스 선택 후, '주문저장' 클릭
                driver.find_element(By.NAME, 'j_id0:mainFrm:multipleOrder:outMultipleOrderList:%d:j_id81' % int(current_pos)).click()
                wait.until(lambda d : driver.find_element(By.ID, 'j_id0:mainFrm:multipleOrder:j_id65:j_id66:j_id67:submitBtn').click() or True)
                wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'blockPage'))) # 로딩창이 뜰때까지 대기 (로딩창 뜨는 속도보다 코드 실행되는 속도가 더 빨라서 멈춰줘야함)
                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'blockPage'))) # 로딩창 닫히면 재개
    if not error:
        chkERR()

def chkERR():
    global error, current_url, current_pos, selected
    if error == 1:
        notify('작업을 중지합니다.',
              '주문번호가 일치하지 않는 항목이 있습니다.\n주문번호 : %s\n확인된 위치 : %d번째 세트항목\n콘솔창에서 다시 시작할 수 있습니다.' % (selected, current_pos+1), 
              scenario='incomingCall')
        print("[에러] 작업을 중지합니다.")
        print("[에러] 주문번호가 일치하지 않는 항목이 있습니다.")
        print("[에러] %d번째 칸의 주문번호가 %s(이)가 아닙니다." % (current_pos+1, selected))
        print("")
        print("")
        input("[정보] Enter를 눌러 order.xlsx를 새로 불러와서 다시 시작할 수 있습니다. (Enter를 눌러 계속)")
        process()
    elif error == 2:
        notify('작업을 중지합니다.', '주문번호가 Error인 항목입니다.\n확인된 위치 : %d번째 세트항목\n콘솔창에서 다시 시작할 수 있습니다.' % (current_pos+1), 
              scenario='incomingCall')
        print("[에러] 작업을 중지합니다.")
        print("[에러] 주문번호가 ERROR인 항목입니다.")
        print("[에러] 현재 URL :", current_url)
        print("[에러] 에러 위치 :", current_pos+1,"번째 세트항목")
        
        print("")
        print("")
        input("[정보] Enter를 눌러 order.xlsx를 새로 불러와서 다시 시작할 수 있습니다. (Enter를 눌러 계속)")
        process()
    
    elif error == 3:
        notify('작업을 중지합니다.', '현재 URL에 아무런 내용이 없습니다.\n해당 URL: %s\n콘솔창에서 다시 시작할 수 있습니다.' % (current_url), 
              scenario='incomingCall')
        print("[에러] 현재 URL의 아무런 내용이 없습니다.")
        print("[에러] order.xlsx의 URL을 다시 한번 확인해 주시고 다시 실행해주세요.")
        print("[에러] 현재 URL :", current_url)

        print("")
        print("")
        input("[정보] Enter를 눌러 order.xlsx를 새로 불러와서 다시 시작할 수 있습니다. (Enter를 눌러 계속)")
        process()
    
    else:
        toast('모든 작업이 완료되었습니다.\n크롬창과 콘솔창을 닫으셔도 됩니다.')
        print('[정보] 모든 작업이 완료되었습니다.')
        print('[정보] 크롬창과 콘솔창을 닫으셔도 됩니다.')

process()


    

    


# In[ ]:





# In[ ]:





# In[ ]:




