#!/usr/bin/env python
# coding: utf-8

# In[1]:


판매처코드 = input("[정보] 판매처 코드를 입력해주세요 : ")
모델번호 = input('[정보] 모델번호를 입력해주세요 : ')


# In[2]:


import selenium
from selenium import webdriver
from selenium.webdriver import ActionChains

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

import os
from datetime import datetime, timedelta


# In[3]:


dates = [] # [['2024. 1. 1', '2024. 1. 3'],['2024. 1. 4', '2024. 1. 6'], ...]

print("[정보] 다운로드 받을 날짜 구간을 입력해주세요.")
print("[정보] 입력 예시)240321-240322")
print("[정보] 입력 예시)240323-240326")
print("[정보] 입력 예시)end")
while True:
    try:
        inp = input("")
        dates.append([datetime.strptime(inp.split("-")[0], '%y%m%d').strftime("%Y. %#m. %#d"), 
                      datetime.strptime(inp.split("-")[1], '%y%m%d').strftime("%Y. %#m. %#d")])
    except:
        print("[에러] 날짜를 형식에 맞게 입력해주셔야 합니다! 방금 입력하신 내용은 제외됐으니 다시 입력해주세요! 모든 입력이 완료되면 'end'를 입력해주세요.")
    if inp == "end":
        break

# start_date = datetime(int(input("시작 연도를 4자리를 입력해주세요 (예)2023 : ")),
#                       int(input("시작 월을 입력해주세요 (예)3 : ")),
#                       int(input("시자 일을 입력해주세요 (예)21 : ")))

# .strftime("%Y. %#m. %#d")

# end_date = datetime(int(input("종료 연도를 4자리를 입력해주세요 (예)2023 : ")),
#                       int(input("종료 월을 입력해주세요 (예)3 : ")),
#                       int(input("종료 일을 입력해주세요 (예)21 : ")))

print("===== [실행 전 확인] =====")


print("판매처코드 :", 판매처코드)
print("모델번호 :", 모델번호)
print("-- 구간 --")
cnt = 1
for date in dates:
    print(cnt,"번째 다운로드 구간 :", date[0], "~", date[1])
    cnt+=1
print("-- -- --")
print("===== ===== ===== =====")
input("[정보] 잘못 입력한 값이 있으면 종료 후 다시 입력해주세요. Enter를 누르면 창이 열립니다.")


# In[4]:


dir = './download'

if not os.path.exists(dir):
    os.makedirs(dir)


# In[5]:


# OLD CODE,직접 날짜, 구간을 입력해야하므로 DEPRECATED
#dates = [] # [['2024. 1. 1', '2024. 1. 3'],['2024. 1. 4', '2024. 1. 6'], ...]
#least_date = start_date

#while least_date < end_date:
#    dates.append([least_date.strftime("%Y. %#m. %#d"), (least_date+timedelta(1)).strftime("%Y. %#m. %#d")])
#    least_date = least_date+timedelta(2)


# In[6]:


service = webdriver.ChromeService(executable_path='chromedriver/chromedriver.exe')

options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : os.path.abspath(dir)}
options.add_experimental_option('prefs', prefs)

driver = webdriver.Chrome(service=service, options=options)


# In[7]:


url = "https://sec-b2b--c.vf.force.com/apex/OrderProgressStatusList?sfdc.tabName=01r28000000ox8N" # 주문 조회페이지로 바로 접속
driver.get(url)
input("[정보] 로그인이 되었으면 Enter를 눌러주세요")


# In[8]:


driver.get(url)
input("[정보] 공지창이 열려있다면 7일 동안 다시보지 않기 버튼을 꼭 눌러주세요! (Enter키를 눌러 계속)")
print("[경고!] 지금부터 자동으로 다운로드가 진행됩니다. 절대 해당 크롬 창을 클릭하지 마세요") 
print("[경고!] 해당 크롬 창은 무시하시고, 다른 작업은 계속 진행하셔도 됩니다.") 
for date in dates:
    driver.switch_to.window(driver.window_handles[0])
    driver.find_element(By.ID, 'j_id0:pageBody:pageForm:j_id39:j_id53:j_id54:j_id57:RequiredField:soldtoinfo_lkwgt').click()
    
    # 판매처 조회칸에 값 입력 (가장 마지막에 열린창)
    driver.switch_to.window(driver.window_handles[-1])
    driver.switch_to.frame('searchFrame')
    driver.find_element(By.NAME, 'lksrch').send_keys(판매처코드)
    driver.find_element(By.ID, 'lkenhmdSEARCH_ALL').click()
    driver.find_element(By.NAME, 'lksrch').send_keys(Keys.ENTER)
    
    driver.switch_to.window(driver.window_handles[-1])
    driver.switch_to.frame('resultsFrame')
    
    driver.find_element(By.CLASS_NAME, 'dataCell').click()
    
    ###
    
    driver.switch_to.window(driver.window_handles[0])
    driver.find_element(By.ID, 'j_id0:pageBody:pageForm:j_id39:j_id53:j_id73:j_id75_lkwgt').click()
    
    # 물품 조회칸에 값 입력 (가장 마지막에 열린창)
    driver.switch_to.window(driver.window_handles[-1])
    driver.switch_to.frame('searchFrame')
    driver.find_element(By.NAME, 'lksrch').send_keys(모델번호)
    driver.find_element(By.NAME, 'lksrch').send_keys(Keys.ENTER)
    
    driver.switch_to.window(driver.window_handles[-1])
    driver.switch_to.frame('resultsFrame')
    
    driver.find_element(By.CLASS_NAME, 'dataCell').click()
    
    driver.switch_to.window(driver.window_handles[0])
    
    driver.find_elements(By.CLASS_NAME, 'rt')[0].clear()
    driver.find_elements(By.CLASS_NAME, 'rt')[0].send_keys(date[0])
    
    driver.find_elements(By.ID, 'j_id0:pageBody:pageForm:j_id39:j_id53:j_id90:j_id93:RequiredField:j_id104:j_id105:j_id108')[0].clear()
    driver.find_elements(By.ID, 'j_id0:pageBody:pageForm:j_id39:j_id53:j_id90:j_id93:RequiredField:j_id104:j_id105:j_id108')[0].send_keys(date[1])
    driver.find_elements(By.ID, 'j_id0:pageBody:pageForm:j_id39:j_id53:j_id90:j_id93:RequiredField:j_id104:j_id105:j_id108')[0].send_keys(Keys.ENTER)

    wait = WebDriverWait(driver, 60)
    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'blockPage'))) # 로딩창이 뜰때까지 대기 (로딩창 뜨는 속도보다 코드 실행되는 속도가 더 빨라서 멈춰줘야함)
    wait = WebDriverWait(driver, 60)
    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, 'blockPage')))
    wait = WebDriverWait(driver, 60) # 원래 wait 쓸때마다 초기화 해줘야하는지 모르겠음. 근데 없으면 무시하고 넘어가서 추가해둠
    wait.until(lambda d : driver.find_element(By.ID, 'j_id0:pageBody:pageForm:j_id39:j_id141:j_id142:xyz').click() or True)


# In[9]:


driver.switch_to.window(driver.window_handles[0])
driver.find_elements(By.ID, 'j_id0:pageBody:pageForm:j_id39:j_id141:j_id142:xyz')[0].click()


# In[10]:


# 마지막 1개 파일의 다운로드 버튼을 누르기 전에 웹이 꺼질 수 있음 보완 필요
input("[정보] 다운로드가 완료된 것으로 확인됩니다. download 폴더를 확인하신 후 Enter를 눌러 종료하세요.")
driver.quit()


# In[ ]:




