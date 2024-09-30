#!/usr/bin/env python
# coding: utf-8

# In[6]:


import os
import xlwings as xw
import pandas as pd
from xlwings.constants import DeleteShiftDirection


# In[7]:


dir = './merge'
result_arr = []


# In[8]:


if not os.path.exists(dir):
    os.makedirs(dir)

input("[정보] merge 폴더에 합칠 엑셀 파일들만 넣어주세요. (Enter를 눌러 계속)")


# In[9]:


files = os.listdir(dir)
files.sort(key = lambda x : os.path.getmtime(dir + "/" + x))


# In[10]:


for file in files:
    book = xw.Book(dir + "/" + file)
    sheet = book.sheets[0]
    sheet.range('1:1').api.Delete(DeleteShiftDirection.xlShiftUp) 
    target = sheet.used_range.options(pd.DataFrame, index=False).value
    book.close()
    
    # 2줄로 된 제목을 위해 제목 변경 ( 하드코딩 수정 예정 )
    target.columns.values[9] = '인수거부'
    target.columns.values[10] = '불량'
    target.drop([0], axis=0, inplace=True)

    result_arr.append(target)

output_target = target.head(0)
output_target = pd.concat(result_arr)
output_target.to_excel("합본.xlsx", engine="openpyxl", index=False)
input("[정보] Merge가 완료되었습니다. (Enter키를 눌러 종료)")

# DUMMY ( 셀 제목 2줄 버전 / 이상하게 작동 )
#for file in os.listdir(dir):
#    book = xw.Book(dir + "/" + file)
#    sheet = book.sheets[0]
#    target = sheet.used_range.options(pd.DataFrame, header=2, index=False).value
#    book.close()
#
#    result_arr.append(target)
#
#output_target = target.head(0)
#output_target = pd.concat([output_target, result_arr])
#output_target.to_excel("합본.xlsx", engine="openpyxl")