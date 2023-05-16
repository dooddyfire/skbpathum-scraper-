import selenium 
import datetime
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
#from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
#Fix
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup

# set options to be headless, ..
from selenium import webdriver
    
import webdriver_manager
import bs4 

import requests 

driver = webdriver.Chrome(ChromeDriverManager().install())

title_lis = []
price_lis = []
url_lis = []
img_lis = []



for page in range(1,19): 
    url = "https://www.skbpathum.com/th/category/126712/all-product?page={}&sort=".format(page)
    driver.get(url)
    soup = BeautifulSoup(driver.page_source,'html.parser')

    for item in soup.find_all('figure',{'itemscope':'itemscope'}):

    #title 
        title = item.find('a')['title']
        title_lis.append(title)
        print(title)

        # link
        link = item.find('a')['href']
        url_lis.append(link)
        print(link)

        img = item.find('img')['src']
        img_lis.append(img)
        print(img)


        price = item.find('div',{'class':'list-price'}).text
        price_lis.append(price.strip())
        print(price.strip())

df = pd.DataFrame()
df['ชื่อสินค้า'] = title_lis 
df['ราคาสินค้า'] = price_lis 
df['รูปสินค้า'] = img_lis 
df['ลิงค์สินค้า'] = url_lis 
df.to_excel("pathum.xlsx")