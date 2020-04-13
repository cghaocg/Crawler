import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from openpyxl import Workbook
import time as tm
from time import time, strftime, localtime
import random

g_li_columnName = ['ISIN Code', 'Issuer Name', 'Security Type', 'FX', 'Term']

def check_target_url(soup, li_columnName):
    # 檢查是否為目標頁面
    is_target = False
    if len(soup.find_all('th')) == 5:
        num = 0
        is_target = True
        
        for col in soup.find_all('th'):
            if col.text.strip() != li_columnName[num]:
                is_target = False
                break
            num += 1
            
    return is_target


def list_and_crawl(host, path):
    #tm.sleep(random.uniform(1.5, 4)) # 停頓，防止網頁阻擋
    resp = requests.get(host + path)
    
    if resp.status_code == requests.codes.ok:
        soup = BeautifulSoup(resp.text, 'lxml')
        
        if check_target_url(soup, g_li_columnName):
        # 如果為目標頁：進行爬取
            print("目標頁面，進行爬取：[" + resp.url + "]")
            
            num = 0
            for data in soup.find_all('td'):
                num += 1
                remain = num % 5  # 取餘數
                
                data = data.text.strip()

                if remain == 1:
                    columnA = data
                elif remain == 2:
                    columnB = data
                elif remain == 3:
                    columnC = data
                elif remain == 4:
                    columnD = data
                elif remain == 0:
                    columnE = data
                    sheet.append([columnA, columnB, columnC, columnD, columnE])

            print("完成爬取頁面：[" + resp.url + "]")

        else:
            if soup.find("h2").text == "Networked International Securities Identification Number Data Record":
            # 如果為獨立頁：進行爬取
                print("獨立頁面，進行爬取：[" + resp.url + "]")
            
                num = 0
                for data in soup.find_all('td'):
                    num += 1

                    data = data.text.strip()

                    if num == 2:
                        columnA = data
                    elif num == 4:
                        columnD = data
                    elif num == 6:
                        columnC = data
                    elif num == 8:
                        columnE = data
                    elif num == 16:
                        columnB = data
                        sheet.append([columnA, columnB, columnC, columnD, columnE])
                        
                print("完成爬取頁面：[" + resp.url + "]")
            
            else:
            # 如果為列表頁：迴圈跑列表的 url，並呼叫自己，往下一頁
                us_url = soup.select('a[href^="/ISIN/prefix/US"]')
                for u in us_url:
                    print("處理 url：" + u.get('href'))
                    list_and_crawl(host, u.get('href'))
                
    else:
        print("response code is NOT OK, code:[" + resp.status_code + "]")


#====================== 匯出 excel 設定

# 用python建立一個Excel空白活頁簿
excel_file = Workbook()

# 建立一個工作表
sheet = excel_file.active

# 先填入第一列的欄位名稱
sheet['A1'] = 'ISIN Code'
sheet['B1'] = 'Issuer Name'
sheet['C1'] = 'Security Type'
sheet['D1'] = 'FX'
sheet['E1'] = 'Term'


print("開始爬取 [http://www.isinlei.com/ISIN/prefix/US]，開始時間：", strftime("%Y-%m-%d %H:%M:%S", localtime()))
t0 = time()
list_and_crawl("http://www.isinlei.com", "/ISIN/prefix/US") #http://www.isinlei.com/ISIN/prefix/US
print('花費時間： %0.3fs' % (time() - t0), strftime("%Y-%m-%d %H:%M:%S", localtime()))

excel_file.save('prefix_US.xlsx')