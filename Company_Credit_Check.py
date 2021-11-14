
import re
import os
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from fake_useragent import UserAgent


## selenium method
def get_soup(number):
    
    '''
    number = '86687633'
    '''

    url = 'https://findbiz.nat.gov.tw/fts/query/QueryList/queryList.do'

    # 加入headless模式 => 注意：使用時google版本可能不一樣
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")

    # 設定fake_agent
    ua = UserAgent()
    options.add_argument(ua.random)

    # 將參數帶入
    driver = webdriver.Chrome(options=options, executable_path="/Users/tim/Desktop/For Github/Company_Credit_Check/chromedriver")
    driver.get(url)

    # 選取所有資料種類 => 第一個為預設值，已打勾，他們的x_path格式都相同，只差在數字上差異
    for n in range(3, 10, 2):
        x_path = f'//*[@id="queryListForm"]/div[1]/div[1]/div/div[4]/div[2]/div/div/div/input[{n}]'
        driver.find_element_by_xpath(x_path).click()
        time.sleep(1)


    # 輸入統一編號 => number
    driver.find_element_by_tag_name("input").send_keys(number)
    time.sleep(3)

    # 點擊查詢
    driver.find_element_by_xpath('//*[@id="qryBtn"]').click()
    time.sleep(3)

    # 過濾結果
    # revised requirement => class = panel panel-default
    # 1. 是否有海外分公司，兩者保存
    # 2. 剔除工廠資料，只留公司即可
    
    search_results = BeautifulSoup(driver.page_source, features='html.parser').select('div[class="panel panel-default"]')
    for seq, result in enumerate(search_results, 1):
        company_type= re.findall('資料種類.*', result.text)[0].split('：')[1]
        if company_type in ['公司', '分公司']: 
            # 只存取公司資料
            if company_type == '公司':
                authorization = re.finadll('登記現況.*', result.text)[0]
                
            driver.find_element_by_xpath(f'//*[@id="vParagraph"]/div[{seq}]/div[1]/a').click()

            # make screenshot
            driver.save_screenshot(f'{number}.png') # 檔案名稱 => 若為總公司＆分公司，則名稱為c-1、c-2
        
            # 回上一頁
            driver.execute_script("window.history.go(-1)")
            

    # get soup
    # soup = BeautifulSoup(driver.page_source, features='html.parser')

    return authorization


# 取得table & 進行資料處理
def extract_table(soup):

    raw_data = soup.select('table')[0]
    df = pd.read_html(str(raw_data))[0]
    
    # data整理
    df = df.T
    df.columns = df.iloc[0, :]
    df_output = df[['統一編號', '公司名稱', '公司狀況']].drop([0])
    df_output = df_output.assign(統一編號 = df_output['統一編號'].apply(lambda x: re.findall('\d+', x)[0]),
                                 公司名稱 = df_output['公司名稱'].apply(lambda x: x.split(' ')[0]),
                                 公司狀況 = df_output['公司狀況'].apply(lambda x: x.split(' ')[0]))

    return df_output

# 讀取統一編號的excel -> 


# move pics to other directory
# def move_screenshot():


def main():
    number_lists = ['11337775', '04200199', '11395000']
    result = pd.DataFrame()
    for num in number_lists:
        soup = get_soup(num)
        df_output = extract_table(soup)
        result = pd.concat([result, df_output]) 
    print(result)   


if __name__ == '__main__':
    main()





'''
# connect to the website
url = 'https://findbiz.nat.gov.tw/fts/query/QueryList/queryList.do'
headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36'}

data = {'qryCond': 23676764}
        # 'isAlive': 'all',
        # 'type': 'cmpyType',
        # 'infoType': 'D'}
        # # 'vtable': False,
        # 'cmpyType': True,
        # 'fhl': 'zh_TW'}
        # # 'JSESSIONID': '839EBBE176DF098C7B7DD3E670924DA7',
        # 'DWRSESSIONID': 'Pcmgo9uRWk8RNfTqdSTNUWPunQn'}

cookies = {'Cookie': 'type=cmpyType%2C; infoType=D; isAlive=all; cityScope=; busiItemMain=; busiItemSub=; vTable=false; surveytimes=7; qryCond=29028891; JSESSIONID=DBB5D9CC7039F78325D1C8735BA266D7; DWRSESSIONID=Pcmgo9uRWk8RNfTqdSTNUWPunQn; _ga=GA1.3.1507810062.1636776337; _gid=GA1.3.1550726567.1636776337; JSESSIONID=64685E3FC89196837D9513522CB04D5A'}


session = requests.session()
res = session.post(url=url, headers=headers, cookies=cookies)

soup = BeautifulSoup(res.text, features='html.parser')

soup.select('div[class="main"]')
'''

