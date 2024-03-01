import time
import yaml
import openpyxl
import re
import pandas as pd
import logging

from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains

config = yaml.load(open('DG4278_config.yml'), Loader=yaml.Loader)
timeout = 20
f = 'DG4278'
time_string = datetime.now().strftime('%Y-%m-%d')
dev_logger: logging.Logger = logging.getLogger(name='dev')
dev_logger.setLevel(logging.DEBUG)

formatter: logging.Formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', datefmt='%Y%m%d %H:%M:%S')

# print on console
handler: logging.StreamHandler = logging.StreamHandler()
handler.setFormatter(formatter)
dev_logger.addHandler(handler)

#save on log
file_handler = logging.StreamHandler(open(f'{time_string}.log', 'w'))
file_handler.setFormatter(formatter)
dev_logger.addHandler(file_handler)

def get_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')                 # 瀏覽器不提供可視化頁面
    options.add_argument('-no-sandbox')               # 以最高權限運行
    options.add_argument('--start-maximized')        # 縮放縮放（全屏窗口）設置元素比較準確
    options.add_argument('--disable-gpu')            # 谷歌文檔說明需要加上這個屬性來規避bug
    options.add_argument('--window-size=1920,1080')  # 設置瀏覽器按鈕（窗口大小）
    options.add_argument('--incognito')               # 啟動無痕

    driver = webdriver.Chrome(options=options)
    url = config['url']

    # driver.implicitly_wait(10)
    # driver.get(url)
    # driver.delete_all_cookies() #清cookie
    
    # with open("cookies.yml", "r") as f:
    #     cookies = yaml.safe_load(f)
    #     for c in cookies:
    #         if 'domain' in c:
    #             c['domain'] = 'xxx'
    #         dev_logger.info(c)
    #         driver.add_cookie(c)

    driver.get(url)
    

    return driver

def complie_data():
    try:
        start = time.time()
        component_title = []
        driver = get_driver()
        login(driver)

        #get components name
        components_page = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sidebar"]/div/div[1]/nav/div/div/ul/li[5]/a')))
        actions = ActionChains(driver)
        actions.click(components_page).perform()
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sidebar-page-container"]/div[1]/div/div/h1')))
        components = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="components-table"]/tbody[2]'))).find_elements(By.TAG_NAME, 'tr')
        for title in components:
            new_titile = title.find_elements(By.TAG_NAME, 'td')[0].text.strip().replace("/", "")
            component_title.append(new_titile)

        component_data = component(driver)

        components_df = pd.DataFrame(component_data)
        with pd.ExcelWriter(f'{time_string}{f}.xlsx', mode="w+", engine="openpyxl") as writer:
            components_df.to_excel(writer, sheet_name='components', index=False)
        
        for i in range(len(components)):
        # for i in range(6,7):
            components_page = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sidebar"]/div/div[1]/nav/div/div/ul/li[5]/a')))
            actions = ActionChains(driver)
            actions.click(components_page).perform()

            check_page = driver.find_element(By.XPATH, '//*[@id="components-table"]/tbody[2]/tr[' + str(i+1) +']/td[1]/div/a').text.replace("/", "")
            # dev_logger.info(check_page, component_title[i])
            if check_page == component_title[i]:
                driver.find_element(By.XPATH, '//*[@id="components-table"]/tbody[2]/tr[' + str(i+1) +']/td[1]/div/a').click()
                dev_logger.info(f'Now at the {check_page}.')
            issus_data = issus(driver)
            dev_logger.info('Go to excel.')
            issus_df = pd.DataFrame(issus_data)
            
            with pd.ExcelWriter(f'{time_string}{f}.xls', mode="a", engine="openpyxl") as writer:
                df = issus_df.fillna('').astype(str)
                for col in df.columns:
                    df[col] = df[col].apply(lambda x: data_clean(x))
                df.to_excel(writer, sheet_name=component_title[i], index=False)
            dev_logger.info(f'{component_title[i]} data appended successfully.')
            jira= WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="logo"]/a')))
            if jira.text == 'CPS Jira':
                jira_text = jira.text
                jira.click()
                dev_logger.info(f'Turn to page {jira_text}.')
        
        end = time.time()
        dev_logger.info('Time elapsed: ' + str(start-end) + ' seconds')

    except exceptions.InvalidCookieDomainException as e:
        dev_logger.critical(e, exc_info=True)

def data_clean(text):
    # 清洗excel中的非法字符，都是不常見的不可顯示字符，例如退格，響鈴等
    ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
    text = ILLEGAL_CHARACTERS_RE.sub(r'', text)
    return text

def login(driver):
    dev_logger.info('Waiting for login...')
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'i0116'))).send_keys(config['username']) 
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'idSIButton9'))).click() 
    time.sleep(3)
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'i0118'))).send_keys(config['password']) 
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'idSIButton9'))).click() 
    number = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'idRichContext_DisplaySign'))).text
    dev_logger.info(f'Please enter number on your phone: {number}')
    time.sleep(8)
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'idSIButton9'))).click() 
    dev_logger.info('Waiting for the website to load...')
    # cookie2 = driver.get_cookies() #取得登入後cookie
    # with open("cookies.yml", "w") as f:
    #     yaml.safe_dump(data=cookie2, stream=f)

def component(driver):
    component_data = {}
    Component_text = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sidebar-page-container"]/div[1]/div/div/h1'))).text
    dev_logger.info(f'Turn to page {Component_text}.')
    if Component_text == 'Components':
        components_table = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="components-table"]/tbody[1]/tr'))).find_elements(By.TAG_NAME, 'th')
        for components_th in components_table:
            component_data[components_th.text]=[]
        
        item_state_ready = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="components-table"]/tbody[2]'))).find_elements(By.TAG_NAME, 'tr')
        for issus_tr in item_state_ready:
            component_data[list(component_data.keys())[0]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[0].text.strip())
            component_data[list(component_data.keys())[1]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[1].text.strip())
            component_data[list(component_data.keys())[2]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[2].text.strip())
            component_data[list(component_data.keys())[3]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[3].text.strip())
            component_data[list(component_data.keys())[4]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[4].text.strip())
            component_data[list(component_data.keys())[5]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[5].text.strip())
            component_data[list(component_data.keys())[6]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[6].text.strip())
    
    return component_data

def issus(driver):
    issus_data = {}
    issus_text = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="search-header-view"]/div/h1'))).text
    dev_logger.info(f'Turn to page {issus_text}.')

    #get value from issus title
    try:
        if issus_text == 'Search':
            time.sleep(1)
            issus_th = driver.find_element(By.XPATH, '//*[@id="issuetable"]/thead/tr').find_elements(By.TAG_NAME, 'th')
            for i in range(1, len(issus_th)):
                issus_th = driver.find_element(By.XPATH, '//*[@id="issuetable"]/thead/tr/th[' + str(i) +']/span').text
                issus_data[issus_th]=[]

    except exceptions.InvalidCookieDomainException as e:
        dev_logger.critical(e, exc_info=True)


    #get value from issus page information
    total = 0
    issus_data['OpCo']=[]
    issus_data['Description']=[]
    issus_data['Link']=[]
    issus_data['Name']=[]
    issus_data['status-lozenge']=[]
    issus_data['last-execution-status']=[]
    # issus_data['play-button']=[]
    # issus_data['remove-button']=[]
    while True:
        tbody = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'tbody'))).find_elements(By.TAG_NAME, 'tr')
        for issus_tr in tbody:
            dev_logger.info(f"Now at {issus_tr.find_elements(By.TAG_NAME, 'td')[1].text.strip()}.")
            WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="search-header-view"]/div/h1')))
            issus_data[list(issus_data.keys())[0]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[0].text.strip())
            issus_data[list(issus_data.keys())[1]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[1].text.strip())
            issus_data[list(issus_data.keys())[2]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[2].text.strip())
            issus_data[list(issus_data.keys())[3]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[3].text.strip())
            issus_data[list(issus_data.keys())[4]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[4].text.strip())
            issus_data[list(issus_data.keys())[5]].append(issus_tr.find_elements(By.TAG_NAME, 'td')[5].text.strip())
            
            #Go to description
            try:
                time.sleep(2)
                issus_tr.find_elements(By.TAG_NAME, 'td')[1].click()

                WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="project-name-val"]')))
                
                #OpCo
                if check_opco(driver) == True:
                    opco_text = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="customfield_24326-field"]/span'))).text.strip()
                    issus_data[list(issus_data.keys())[6]].append(opco_text)
                else:
                    issus_data[list(issus_data.keys())[6]].append('')

                #Description
                key_text = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="key-val"]'))).text.strip()
                summary_text = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="summary-val"]'))).text.strip()

                if check_description(driver) == True:
                    description_text = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="descriptionmodule"]/div[2]'))).text.strip()
                    # dev_logger.info(description_text)
                    if key_text == list(issus_data.values())[1][-1] and summary_text == list(issus_data.values())[2][-1]:
                        issus_data[list(issus_data.keys())[7]].append(description_text)
                else:
                    issus_data[list(issus_data.keys())[7]].append('')
            
            except exceptions.InvalidCookieDomainException as e:
                dev_logger.critical(e, exc_info=True)

            #Traceability
            try:
                no_textcase = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ZephyrScaleIssuePanel"]/span/section/main/div/span[2]'))).text 
                if 'No test cases.' in no_textcase:
                    # pass
                    issus_data[list(issus_data.keys())[8]].append('')
                    issus_data[list(issus_data.keys())[9]].append('')
                    issus_data[list(issus_data.keys())[10]].append('')
                    issus_data[list(issus_data.keys())[11]].append('')
                    # issus_data[list(issus_data.keys())[12]].append('')
            except:
                total_div = []
                if load_more(driver) == True:
                    total_li = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, 'css-9duhdc'))).find_elements(By.TAG_NAME, 'div')
                    for i in total_li:
                        total_div.append(i.text)

                    load_check = (int((total_div[-1][-2:].strip()))-1)/5   #暫時用-1的方式，以達除完有餘數
                    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ZephyrScaleIssuePanel"]/span/section/main/h2')))
                    for click in range(int(load_check)):
                        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, 'css-rvtbkj'))).click()
                        time.sleep(3)

                textcase = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, 'css-afeyfj'))).find_elements(By.TAG_NAME, 'li')
                # dev_logger.info(len(textcase))
                a = 0 #第一行省略
                for issus_span in textcase:
                    # WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, 'span')))
                    # dev_logger.info(issus_span.find_elements(By.TAG_NAME, 'span')[6].get_attribute('aria-label'))
                    if a != 0:
                        issus_data[list(issus_data.keys())[0]].append('')
                        issus_data[list(issus_data.keys())[1]].append('')
                        issus_data[list(issus_data.keys())[2]].append('')
                        issus_data[list(issus_data.keys())[3]].append('')
                        issus_data[list(issus_data.keys())[4]].append('')
                        issus_data[list(issus_data.keys())[5]].append('')
                        issus_data[list(issus_data.keys())[6]].append('')
                        issus_data[list(issus_data.keys())[7]].append('')

                    issus_data[list(issus_data.keys())[8]].append(issus_span.find_elements(By.TAG_NAME, 'span')[0].text.strip())
                    issus_data[list(issus_data.keys())[9]].append(issus_span.find_elements(By.TAG_NAME, 'span')[1].text.strip())
                    issus_data[list(issus_data.keys())[10]].append(issus_span.find_elements(By.TAG_NAME, 'span')[2].text.strip())
                    issus_data[list(issus_data.keys())[11]].append(issus_span.find_elements(By.TAG_NAME, 'div')[3].get_attribute('aria-label'))
                    # issus_data[list(issus_data.keys())[12]].append(issus_span.find_elements(By.TAG_NAME, 'span')[6].get_attribute('aria-label'))
                    # issus_data[list(issus_data.keys())[13]].append(issus_span.find_elements(By.TAG_NAME, 'span')[8].get_attribute('aria-label'))

                    a += 1
           
            dev_logger.info(f'Description and Test Cases has been added to {key_text}.')
            driver.back()
            time.sleep(2)

        # for key, value in issus_data.items():
        #     dev_logger.info(key, len([item for item in value if item]))

        total+=(len(tbody))
        results = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div/div[2]/div/div/div/div[2]/div[1]/span'))).text
        if next_iselement(driver) == True:
            next = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.CLASS_NAME, 'nav-next')))
            actions = ActionChains(driver)
            actions.click(next).perform()

            check_title = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div/div[2]/div/div/div/div[2]/div[2]/div/strong'))).text
            dev_logger.info(results, f', Currently on page {check_title}.')

        else:
            dev_logger.info(results)

        # dev_logger.info(total)
        time.sleep(2)
        check_results = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/div/div[2]/div/div/div/div[2]/div[1]/span/span[3]'))).text
        # if (total == 25): 
        if (total == int(check_results)): 
            break
    
    return issus_data

def next_iselement(driver):
    try:
        driver.find_element(By.CLASS_NAME, 'nav-next')
        return True
    except exceptions.NoSuchElementException:
        return False

def load_more(driver):
    try:
        driver.find_element(By.CLASS_NAME, 'css-9duhdc')
        return True
    except exceptions.NoSuchElementException:
        return False
    
def check_description(driver):
    try:
        driver.find_element(By.XPATH, '//*[@id="descriptionmodule-label"]')
        return True
    except exceptions.NoSuchElementException:
        return False
    
def check_opco(driver):
    try:
        driver.find_element(By.XPATH, '//*[@id="rowForcustomfield_24326"]/div/strong/label')
        return True
    except exceptions.NoSuchElementException:
        return False


if __name__ == '__main__':
    complie_data()