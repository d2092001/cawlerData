# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 13:53:02 2022

@author: NM DUC
"""
import numpy as np
from selenium import webdriver
from time import sleep
import random
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.by import By
import pandas as pd
import re
import openpyxl
import requests
from bs4 import BeautifulSoup
import json
import urllib3
import json
import urllib.request



excel = pd.read_excel('./auto_find.xlsx')

options = webdriver.ChromeOptions()
options.headless = True
# OR options.add_argument("--disable-gpu")


# Declare browser
driver = webdriver.Chrome(executable_path='./chromedriver.exe',chrome_options=options)

# lấy dữ liệu trong file excel đầu vào 
links = excel['URL'].values.tolist()
#text_finds = excel['Từ khóa cần tìm'].values.tolist()

count_text = []
urls = []
texts = []


# Open URL
for i in range(0,len(links)):
    
    #lấy url ra để xét    
    url = str(links[i]) 
    
    #lấy text cần tìm trong url ở trên
    #text = str(text_finds[i])
    text = 'amis hrm'
    try:
        print(url)
        req = requests.get(url)
        #req = requests.get('https://asp.misa.vn/')
        print(req.status_code)
    except requests.exceptions.RequestException as e:  # This is the correct syntax
        raise SystemExit(e)

    #lấy toàn bộ nội dung website
    soup = BeautifulSoup(req.text, "lxml")
    
    # đưa url đang duyệt vào danh sách url đã duyệt 
    urls.append(url)
    #sleep(1)

    #kiểm tra trc nếu ko thấy phần tử thì bỏ qua url đó 
    if soup.find("div", {"class":"text-404"}) : 
        
        findCountText = "lỗi 301 "
        
    else:
        #lấy chuẩn hơn
        #lọc lấy hết thẻ p trong td-post-content tagdiv-type (có vẻ có cả heading)
        article_text = ''
        article = soup.find("div", {"class":"td-post-content tagdiv-type"}).find_all('p')
        for element in article:
            article_text += '\n' + ''.join(element.findAll(text = True))
            
        #lấy thêm thẻ ul
        article_text_ul = ''
        article_ul = soup.find("div", {"class":"td-post-content tagdiv-type"}).find_all('ul')
        for element_ul in article_ul:
            article_text_ul += '\n' + ''.join(element_ul.findAll(text = True))
        
        
        #lấy thêm thẻ h2
        article_text_h2 = ''
        article_h2 = soup.find("div", {"class":"td-post-content tagdiv-type"}).find_all('h2')
        for element_h2 in article_h2:
            article_text_h2 += '\n' + ''.join(element_h2.findAll(text = True))
            
            
        #lấy tiêu đề
        titleFind = soup.find_all('div', class_='current-title')[0].text
        
        #tổng text
        allTextFind = titleFind + article_text + article_text_ul + article_text_h2
        
        #check xem số lượng bao nhiêu
        findCountText = allTextFind.lower().count(text.lower())
        
    #tìm xong thì đưa vào text
    texts.append(findCountText)

print(urls)

# gép 2 mảng lại để thành file
df1 = pd.DataFrame(list(zip(urls, texts)), columns = ['title', 'số lần hiện'])

# đưa df ra file excel
df1.to_excel("output.xlsx", sheet_name='Sheet_name_1')
print(df1)




# Close browser
driver.close()   
