
import re
import os
import shutil
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from fake_useragent import UserAgent


def search_result_check(driver, company):
    # 過濾結果
    # revised requirement => class = panel panel-default
    # 1. 是否有海外分公司，兩者保存
    # 2. 剔除工廠資料，只留公司即可
    search_results = BeautifulSoup(driver.page_source, features='html.parser').select('div[class="panel panel-default"]')
    n = 1
    authorization = None
    for seq, result in enumerate(search_results, 1): # 從1開始enumerate
        company_type= re.findall('資料種類.*', result.text)[0].split('：')[1]
        if company_type in ['公司', '分公司']: 
            # 只存取公司登記現況資料
            if company_type == '公司':
                authorization = re.findall('登記現況.*', result.text)[0].split('：')[1]
                file_name_new = f"{company}.pdf"
            else:
                file_name_new = f"{company}-{n}.pdf"
                n += 1
                
            driver.find_element_by_xpath(f'//*[@id="vParagraph"]/div[{seq}]/div[1]/a').click()

            # download pdf
            # 檔案名稱 => 若為總公司＆分公司，則名稱為c-1、c-2
            driver.execute_script('window.print()')
            time.sleep(2)

            # 轉移下載檔案到目標資料夾
            # move the file to specific directory => 下載後會存到download資料夾裡面(檔名：商工登記公示資料查詢服務.pdf)
            src_folder = r"/Users/tim/Downloads/"
            dst_folder = r"/Users/tim/Desktop/selenium_download/"
            file_name = '商工登記公示資料查詢服務.pdf'
            shutil.move(src_folder+file_name, dst_folder+file_name_new)
        
            # 回上一頁 => 若只有一頁就不需要
            driver.execute_script("window.history.go(-1)")
            time.sleep(1)
    
    return authorization


## selenium method
def get_authorization_result(company, id):
    
    '''
    id = '22099131'
    company = 'Taiwan Semiconductor Manufacturing Company Limited'
    '''

    url = 'https://findbiz.nat.gov.tw/fts/query/QueryList/queryList.do'

    # 加入headless模式 => 注意：使用時google版本可能不一樣
    options = webdriver.ChromeOptions()

    # prefs = {'profile.default_content_settings.popups': 0,
    #          'download.default_directory': '/Users/tim/Desktop/selenium_download'}
    # options.add_experimental_option('prefs', prefs)
    options.add_argument('--kiosk-printing') # for downloading pdf without click any button
    options.add_argument("--headless")

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
    driver.find_element_by_tag_name("input").send_keys(id)
    time.sleep(3)

    # 點擊查詢
    driver.find_element_by_xpath('//*[@id="qryBtn"]').click()
    time.sleep(3)

    # 翻頁需求 => 進到查詢結果頁，判斷有多少個頁面(如果只有一個，則直接執行，若有數個則for loop執行)
    pages_amount = BeautifulSoup(driver.page_source, features='html.parser').select('ul[class="pagination"] li')
    if pages_amount:
        k = 2
        for _ in range(len(pages_amount) - 1):
            # 預設第一頁 => 點擊第一頁
            driver.find_element_by_xpath(f'//*[@id="QueryList_queryList"]/div/div/div/div/nav/ul/li[{k}]/a').click() # 點擊分頁 => 第一頁2、第二頁3

            # 點擊後判斷是否有公司or分公司
            data = search_result_check(driver, company)
            if data != None:
                authorization = data

            # 準備往下一頁移動
            k += 1

    else:
        authorization = search_result_check(driver, company)
    
    driver.close()

    return authorization
        


def main():
    df = pd.read_excel('/Users/tim/Desktop/自動化查詢工商1115.xlsx', sheet_name='temp')
    company_list = df['A/c name'].tolist()
    id_list = df['ID'].tolist()
    auth_data = [] 
    for company, id in zip(company_list, id_list):
        print(company, id)
        try:
            authorization = get_authorization_result(company, id) 
            auth_data.append(authorization)
            print(auth_data)
            time.sleep(1)
        except:
            auth_data.append(None)

    # 查詢結果設為欄位
    df['result'] = auth_data
    
    # show final result
    print(df)

    # 存入原本的excel，作為output sheet
    with pd.ExcelWriter('/Users/tim/Desktop/自動化查詢工商_output.xlsx', engine='openpyxl', mode='a') as writer:  
        df.to_excel(writer, sheet_name='result_1121')

    

if __name__ == '__main__':
    main()




'''
Developing Note:

2021.11.15
1. error handling
2. save as pdf & headquarter and branch name conditions 
3. append sheet on existing excel file => fixed! specify engine as openpyxl!

11/16 update
* use A/C name as the filename and os package to move the file to right directory
* 若查詢結果過多，需進行翻頁 ex. 台積電！

11/20 update
* 面對multi-pages時，已設定好邏輯 => 運行順利
* 無效統編的error handling => OK

11/21 update
*核准設立部分 -> fixed

'''



