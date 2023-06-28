import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

s_list = []


def get_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')                 # 瀏覽器不提供可視化頁面
    options.add_argument('-no-sandbox')               # 以最高權限運行
    options.add_argument('--start-maximized')        # 縮放縮放（全屏窗口）設置元素比較準確
    options.add_argument('--disable-gpu')            # 谷歌文檔說明需要加上這個屬性來規避bug
    options.add_argument('--window-size=1920,1080')  # 設置瀏覽器按鈕（窗口大小）
    options.add_argument('--incognito')               # 啟動無痕

    driver = webdriver.Chrome(options=options)
    driver.get('https://tw.stock.yahoo.com/quote/2330.TW/technical-analysis')

    return driver


def stock():
    driver = get_driver()

    s = driver.find_element(
        By.XPATH, '//*[@id="main-0-QuoteHeader-Proxy"]/div/div[1]')
    s_list.append(s.text.split())

    driver.find_element(By.XPATH, '//*[@id="subscription-banner"]')
    if True:
        driver.find_element(
            By.XPATH, '//*[@id="subscription-banner"]/div/button').click()

    driver.switch_to.frame(0)
    select_element = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, 'TAChartPeriod')))

    if select_element:
        t = select_element.find_element(
            By.XPATH, '//*[@id="TAChartPeriod"]/option[6]')
        t.click()
        element = driver.find_element(By.XPATH, '//*[@id="TaDivBase"]')
        element.screenshot(s_list[0][0] + '_' + s_list[0][1] + '.png')


stock()
