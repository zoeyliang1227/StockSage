import time
import pandas as pd
import re

from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common import exceptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

timeout=10
time_string = datetime.now().strftime('%Y-%m-%d')
stock_data = {}
stock_data['股票名稱']=[]
stock_data['股票代號']=[]

def get_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')                 # 瀏覽器不提供可視化頁面
    options.add_argument('-no-sandbox')               # 以最高權限運行
    options.add_argument('--start-maximized')        # 縮放縮放（全屏窗口）設置元素比較準確
    options.add_argument('--disable-gpu')            # 谷歌文檔說明需要加上這個屬性來規避bug
    options.add_argument('--window-size=1920,1080')  # 設置瀏覽器按鈕（窗口大小）
    options.add_argument('--incognito')               # 啟動無痕
    options.add_experimental_option('prefs', {'profile.managed_default_content_settings.images': 2})      #不顯示任何圖片


    driver = webdriver.Chrome(options=options)
    driver.get('https://tw.stock.yahoo.com/class-quote?sectorId=26&exchange=TAI')

    return driver


def stock():
    try:
        start = time.time()
        driver = get_driver()
        # check_total_length(driver)

        stock_total = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.ID, 'main-1-ClassQuotesTable-Proxy'))).find_elements(By.TAG_NAME, 'li')

        # print(len(stock_total))
        total = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.ID, 'main-1-ClassQuotesTable-Proxy'))).text[2:5]
        for i in range(int(total)+1):   
        # for i in range(175):
        # for i in range(175, int(total)):
            print('開始爬股票')
            division = i//30
            if i >= len(stock_total):
                length = 1440*division
                driver.execute_script(f"window.scrollTo(0, {length})")
                time.sleep(2)
                print(f'下拉 {length}')
            last = int(total) - i
            stock = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-1-ClassQuotesTable-Proxy"]/div/div/div[3]/div[2]/div/div/ul/li[' + str(i+1) +']/div/div[1]/div[2]/div/div[1]'))).text
            code = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-1-ClassQuotesTable-Proxy"]/div/div/div[3]/div[2]/div/div/ul/li[' + str(i+1) +']/div/div[1]/div[2]/div/div[2]'))).text
            print(f'{i}, {stock}, {code}, 目前還剩{last}')
            stock_data[list(stock_data.keys())[0]].append(stock)                                
            stock_data[list(stock_data.keys())[1]].append(code) 

            #開啟新分頁
            href = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-1-ClassQuotesTable-Proxy"]/div/div/div[3]/div[2]/div/div/ul/li[' + str(i+1) +']/div/div[1]/div[2]/div/a'))).get_attribute('href')
            driver.execute_script("window.open('');")
            driver.switch_to.window(driver.window_handles[1]) 
            driver.get(href)
            # WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-1-ClassQuotesTable-Proxy"]/div/div/div[3]/div[2]/div/div/ul/li[' + str(i+1) +']/div/div[1]/div[2]/div/a'))).click()
            # if check_ad(driver) == True:
            #     WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="subscription-banner"]/div/button'))).click()

            time.sleep(2)
            check_stock = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-0-QuoteHeader-Proxy"]/div/div[1]/h1'))).text
            print(f'開啟新網頁，進到 {check_stock} 頁面')
            if check_stock == stock:
                payment(driver)
                print(f'{stock}：股利頁面資料取得完成')

            #關掉新分頁，回到原本的頁面
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

    except exceptions.InvalidCookieDomainException as e:
        print(e, exc_info=True)

    # print(stock_data)
    write_to_excel()
    end = time.time()
    print('Time elapsed: ' + str(start-end) + ' seconds')

def write_to_excel():
    for key in list(stock_data.keys()):
        if not stock_data.get(key):
            stock_data.pop(key)
    print(list(stock_data.keys()))
    for d in stock_data[list(stock_data.keys())[3]]:
        try:
            d = float(d)
        except ValueError:
            d = ''
    # for key, value in stock_data.items():
    #     print(key, len([item for item in value if item]))
    stock_df = pd.DataFrame(stock_data)
    with pd.ExcelWriter(f'{time_string}.xls', mode="w+", engine="openpyxl") as writer:
        stock_df.to_excel(writer, sheet_name='stock', index=False)

def payment(driver):
    if check_payment(driver) == True:     
        WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-1-QuoteTabs-Proxy"]/nav/div/div/div[5]/a'))).click()
                                                                                    

    if len(list(stock_data.keys())) == 2:
        payment = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-2-QuoteDividend-Proxy"]/div/section[2]/div[3]/div[1]/div'))).find_elements(By.TAG_NAME, 'div')
        # print(len(payment))
        
        #股利頁面
        for pay in payment:
            # print(pay.text.strip())
            stock_data[pay.text.strip()]=[]
    if check_dividend(driver) == False:
        stock_data[list(stock_data.keys())[2]].append('')           
        stock_data[list(stock_data.keys())[5]].append('')     
    else:
        dividend_total = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="main-2-QuoteDividend-Proxy"]/div/section[2]/div[3]/div[2]/div/div/ul'))).find_elements(By.TAG_NAME, 'li')
        # print(len(dividend_total))
        a = 0 #第一行省略
        implement_time = 0
        for dividend in dividend_total:
            # str.isdigit() 」可以判斷字串中是否都是數字( 不能包含英文、空白或符號)，回傳True 或False
            if dividend.find_elements(By.TAG_NAME, 'div')[2].text.isdigit() == True:
                if a != 0:
                    stock_data[list(stock_data.keys())[0]].append('')                                
                    stock_data[list(stock_data.keys())[1]].append('') 
                    
                # print(dividend.find_elements(By.TAG_NAME, 'div')[2].text.strip(), dividend.find_elements(By.TAG_NAME, 'div')[5].text.strip())
                stock_data[list(stock_data.keys())[2]].append(dividend.find_elements(By.TAG_NAME, 'div')[2].text.strip())
                # stock_data[list(stock_data.keys())[3]].append(dividend.find_elements(By.TAG_NAME, 'div')[3].text.strip())
                # stock_data[list(stock_data.keys())[4]].append(dividend.find_elements(By.TAG_NAME, 'div')[4].text.strip())
                stock_data[list(stock_data.keys())[5]].append(dividend.find_elements(By.TAG_NAME, 'div')[5].text.strip())
                # stock_data[list(stock_data.keys())[6]].append(dividend.find_elements(By.TAG_NAME, 'div')[6].text.strip())
                # stock_data[list(stock_data.keys())[7]].append(dividend.find_elements(By.TAG_NAME, 'div')[7].text.strip())
                # stock_data[list(stock_data.keys())[8]].append(dividend.find_elements(By.TAG_NAME, 'div')[8].text.strip())
                # stock_data[list(stock_data.keys())[9]].append(dividend.find_elements(By.TAG_NAME, 'div')[9].text.strip())
                # stock_data[list(stock_data.keys())[10]].append(dividend.find_elements(By.TAG_NAME, 'div')[10].text.strip())
                # stock_data[list(stock_data.keys())[11]].append(dividend.find_elements(By.TAG_NAME, 'div')[11].text.strip())
                # stock_data[list(stock_data.keys())[12]].append(dividend.find_elements(By.TAG_NAME, 'div')[12].text.strip())
                # stock_data[list(stock_data.keys())[13]].append(dividend.find_elements(By.TAG_NAME, 'div')[13].text.strip())    
                implement_time += 1
            a += 1
            # 執行次數
            if implement_time == 5:
                break  

def check_total_length(driver):
    temp_height=0
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")   #捲軸拖曳到瀏覽器的最下方
        time.sleep(2)
        check_height = driver.execute_script("return document.documentElement.scrollTop || window.pageYOffset || document.body.scrollTop;")
        if check_height == temp_height:
            break
        temp_height = check_height
        print(check_height)

    driver.execute_script("var q=document.documentElement.scrollTop=0")
    time.sleep(2)

def check_dividend(driver):
    try:
        driver.find_element(By.XPATH, '//*[@id="main-2-QuoteDividend-Proxy"]/div/section[2]/div[3]/div[2]/div/div/ul')
        return True
    except exceptions.NoSuchElementException:
        return False
    
def check_ad(driver):
    try:
        driver.find_element(By.ID, 'footer-1-SubscriptionBanner-Proxy')
        return True
    except exceptions.NoSuchElementException:
        return False
    
def check_payment(driver):
    try:
        driver.find_element(By.XPATH, '//*[@id="main-1-QuoteTabs-Proxy"]/nav/div/div/div[5]/a')
        return True
    except exceptions.NoSuchElementException:
        return False

if __name__ == '__main__':
    stock()

