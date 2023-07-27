from bs4 import BeautifulSoup 
import requests 
import pandas as pd 
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
#Fix
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
import pyperclip

import time

url = "https://www.chiangmaiexpert.com/code/hotel2-en.php"
driver = webdriver.Chrome(service= Service(ChromeDriverManager().install()))
driver.maximize_window()

df = pd.read_excel('eng.xlsx')

no_lis = [ int(d) for d in df['ลำดับ'].to_list()]

agoda_lis = [ c for c in df['Link Agoda']]
print(agoda_lis)


print(no_lis)


excel_row = df.values.tolist()
print(excel_row)

driver.get(url)
row_num = 1
for i in excel_row: 
    
    filename = "{}.txt".format(row_num)
    print(filename)

    place_value = i[3]
    print('Place : {}'.format(place_value))
    place = driver.find_element(By.XPATH,'/html/body/form/input[1]')
    place.send_keys(place_value)


    agoda_hid_value = int(i[4])
    print('Hotel ID : ',agoda_hid_value)
    agoda_hid = driver.find_element(By.XPATH,'/html/body/form/input[2]')
    agoda_hid.send_keys(agoda_hid_value)


    url_book_value = i[5]
    print('url book : ',url_book_value)
    url_book = driver.find_element(By.XPATH,'/html/body/form/input[3]')
    url_book.send_keys(url_book_value)




    pic1_value = str(i[6])
    print('pic1 value : ',pic1_value)
    pic1 = driver.find_element(By.XPATH,'//*[@id="user_input1"]')
    pic1.send_keys(pic1_value)


    pic2_value = str(i[7])
    print('pic2 value : ',pic2_value)
    pic2 = driver.find_element(By.XPATH,'//*[@id="user_input2"]')
    pic2.send_keys(pic2_value)


    pic3_value = str(i[8])
    print('pic3 value : ',pic3_value)
    pic3 = driver.find_element(By.XPATH,'//*[@id="user_input3"]')
    pic3.send_keys(pic3_value)


    pic4_value = str(i[9])
    print('pic4 value : ',pic4_value)
    pic4 = driver.find_element(By.XPATH,'//*[@id="user_input4"]')
    pic4.send_keys(pic4_value)

    #star
    star_value = int(i[10])

    if star_value == 1:
        rad = driver.find_element(By.XPATH,'/html/body/form/input[4]')
    elif star_value == 2:
        rad = driver.find_element(By.XPATH,'/html/body/form/input[5]')    
    elif star_value == 3:
        rad = driver.find_element(By.XPATH,'/html/body/form/input[6]')
    elif star_value == 4:
        rad = driver.find_element(By.XPATH,'/html/body/form/input[7]') 
    
    elif star_value == 5: 
        rad = driver.find_element(By.XPATH,'/html/body/form/input[8]') 
        
    rad.click()

    # find id of option
    x = driver.find_element(By.ID,'hotelpoint')
    drop=Select(x)
    
    # score
    select_value = float(i[11])

    #select_value = float('9.8')
    # select by visible text
    if select_value == 9.9:
        drop.select_by_index("1")

    elif select_value == 9.8:
        drop.select_by_index("2")

    elif select_value == 9.7:
        drop.select_by_index("3")

    elif select_value == 9.6:
        drop.select_by_index("4")

    elif select_value == 9.5:
        drop.select_by_index("5")

    elif select_value == 9.4:
        drop.select_by_index("6")


    elif select_value == 9.3:
        drop.select_by_index("7")


    elif select_value == 9.2:
        drop.select_by_index("8")

    elif select_value == 9.1:
        drop.select_by_index("9")

    elif select_value == 9.0:
        drop.select_by_index("10")

    elif select_value == 8.9:
        drop.select_by_index("11")

    elif select_value == 8.8:
        drop.select_by_index("12")

    elif select_value == 8.7:
        drop.select_by_index("13")

    elif select_value == 8.6:
        drop.select_by_index("14")

    elif select_value == 8.5:
        drop.select_by_index("15")

    elif select_value == 8.4:
        drop.select_by_index("16")

    elif select_value == 8.3:
        drop.select_by_index("17")

    elif select_value == 8.2:
        drop.select_by_index("18")

    elif select_value == 8.1:
        drop.select_by_index("19")

    elif select_value == 8.0:
        drop.select_by_index("20")

    elif select_value == 7.9:
        drop.select_by_index("21")

    elif select_value == 7.8:
        drop.select_by_index("22")

    elif select_value == 7.7:
        drop.select_by_index("23")


    elif select_value == 7.6:
        drop.select_by_index("24")

    elif select_value == 7.5:
        drop.select_by_index("25")

    elif select_value == 7.6:
        drop.select_by_index("26")

    elif select_value == 7.5:
        drop.select_by_index("27")


    elif select_value == 7.4:
        drop.select_by_index("28")


    elif select_value == 7.3:
        drop.select_by_index("29")

 
    elif select_value == 7.2:
        drop.select_by_index("30")


    elif select_value == 7.1:
        drop.select_by_index("31")


    elif select_value == 7.0:
        drop.select_by_index("32")



    elif select_value == 6.9:
        drop.select_by_index("33")                       


    elif select_value == 6.8:
        drop.select_by_index("34")     


    elif select_value == 6.7:
        drop.select_by_index("35")     


    elif select_value == 6.6:
        drop.select_by_index("36")


    elif select_value == 6.5:
        drop.select_by_index("37")          

    elif select_value == 6.4:
        drop.select_by_index("38") 

    elif select_value == 6.3:
        drop.select_by_index("39") 

    elif select_value == 6.2:
        drop.select_by_index("40")

    elif select_value == 6.1:
        drop.select_by_index("41")

    elif select_value == 6.0:
        drop.select_by_index("42")   

    start_price_value = '{:,}'.format(int(i[13]))
    start_price = driver.find_element(By.XPATH,'/html/body/form/input[9]')
    start_price.send_keys(start_price_value)

    end_price_value = '{:,}'.format(int(i[14]))
    end_price = driver.find_element(By.XPATH,'/html/body/form/input[10]')
    end_price.send_keys(end_price_value)

    
    has_breakfast = i[15]


    breakfast = driver.find_element(By.XPATH,'/html/body/form/input[11]')
    breakfast.click()

    no_breakfast = driver.find_element(By.XPATH,'/html/body/form/input[12]')


    if has_breakfast == "มี": 
        breakfast.click()
    
    elif has_breakfast == "ไม่มี":
        no_breakfast.click()


    address_value = i[16]
    address = driver.find_element(By.XPATH,'/html/body/form/input[13]')
    address.send_keys(address_value)


    review_value=i[17]
    review = driver.find_element(By.XPATH,'/html/body/form/textarea[5]')
    review.send_keys(review_value)


    check_btn = driver.find_element(By.XPATH,'/html/body/input[2]')
    check_btn.click()
    
    input("Please enter to continue : ")


    
    gen_code = driver.find_element(By.XPATH,'/html/body/form/input[14]')
    gen_code.click()

    copy_code = driver.find_element(By.ID,'copy')
    copy_code.click()

    reset = driver.find_element(By.XPATH,'/html/body/input[1]')




    with open(filename,'w',encoding="utf-8") as f: 
        value_text = pyperclip.paste()
        print(value_text)
        lis_write = value_text.split("\n")
        f.writelines(lis_write)

    reset.click()
    row_num = row_num + 1
    
    #เอาออก
    time.sleep(5)