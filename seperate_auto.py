#!/usr/bin/env python
# coding: utf-8

# In[1]:


print("[정보] 필요 라이브러리를 불러옵니다.")
import numpy as np
import pandas as pd
import math
import xlwings as xw
print("[정보] 필요 라이브러리를 불러왔습니다.")


# In[2]:


# 약 5초 소요
input("[정보] input.xlsx 파일을 불러옵니다. 파일이 준비되어 있으면 Enter키를 눌러주세요.")
book = xw.Book('input.xlsx')
sheet = book.sheets[0]
target = sheet.used_range.options(pd.DataFrame, index=False).value
book.close()
print("[정보] input.xlsx 파일을 불러왔습니다.")


# In[3]:


# 2줄로 된 제목을 위해 제목 변경 ( 하드코딩 수정 예정 )
# target.columns.values[9] = '인수거부'
# target.columns.values[10] = '불량'


# In[4]:


# 반드시 한번만 실행! 2번 이상 실행했을경우 처음부터 실행해야 함 ( 조건문 추가 필요 )
# target.drop([0], axis=0, inplace=True)
# output_target = target.head(0)


# In[3]:


#고객명+생일로 열 추가
target['고객명생일'] = target['고객명'].str.split("_").str[1]+target['고객명'].str.split("_").str[2]


# In[4]:


target = target.replace({np.nan: None})
person_list = target['고객명생일']


# In[5]:


# 고객명생일 형식 아닌거 지우기 (재단명, Null값은 여기서 삭제됨)
person_list = list(filter(None,person_list))


# In[6]:


#중복제거
person_list = list(dict.fromkeys(person_list))


# In[7]:


print("[정보] 분류를 시작합니다")
i = 0
value = 0
arr_df = []

for p in person_list:
    value = 0
    for i in target.loc[target['고객명생일'].str.contains(p, na=False, regex=False)]['주문유형'].tolist():
        if i == 'YKKR-ZFM': #설치 후 고장
            value -= 1
        elif i == 'YKA1-ZZB': #단순변심
            value -= 1
        elif i == 'YKB2-ZZA': #설치계약
            value += 1
        # 1일 경우 일단 설치계약 된걸로 이것만 가져가면 됨!
        # 0과 같거나 보다 작을경우 설치되지 않았으니 걸러야함!
        # 1보다 클 경우 중복주문 된거니 마지막에 주문된걸 살리고 상단에 있는 주문을 걸러야함!
    
    if value >= 1:
        tmp_df = target.loc[target['고객명생일'].str.contains(p, na=False, regex=False) & target['주문유형'].str.contains('YKB2-ZZA', na=False, regex=False),:]
        value = len(tmp_df)
        for i in range(value-1):
            print("[정보]", p, "님의 일부 데이터를 삭제합니다.")
            tmp_df = tmp_df.drop(tmp_df.index[0]) # 맨 위에것만 지우면 가장 마지막인 맨 밑에만 남을테니..
        arr_df.append(tmp_df)

output_target = pd.concat(arr_df)


# In[8]:


print("")
print("")
print("[정보] 정리된 후 남은 데이터 수 :", len(output_target))


# In[9]:


output_target = output_target.sort_index()
output_target = output_target.drop('고객명생일', axis=1)


# In[10]:


output_target.to_excel("output.xlsx", engine="openpyxl", index=False)
input("[정보] output.xlsx 파일을 확인해주세요. (Enter를 눌러 종료)")


# In[ ]:




