from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from pprint import pprint
import numpy as np
import openpyxl
import sys
import random
import selenium.common.exceptions as ec
import os


def getTable(driver):
    main = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "equityStockTable"))
    )
    tbod=main.find_element(By.TAG_NAME,"tbody")
    trs=tbod.find_elements(By.TAG_NAME,"tr")
    print(len(trs))
    for tr in trs:
        tds=tr.find_elements(By.TAG_NAME,"td")
        print(len(tds))
        # for td in tds:
        if(len(tds)!=17):
            continue
        i=0
        for keys, values in col_dict.items(): 
            print(keys)
            # if()
            col_dict[keys].append(tds[i].text)
            i=i+1
        
    pprint(col_dict)
def getOptions(driver):
    main = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "optionChainTable-indices"))
    )
    tbod=main.find_element(By.TAG_NAME,"tbody")
    trs=tbod.find_elements(By.TAG_NAME,"tr")
    print(len(trs))
    for tr in trs:
        tds=tr.find_elements(By.TAG_NAME,"td")
        print(len(tds))
        # for td in tds:
        # if(len(tds)!=17):
        #     continue
        i=1
        for keys, values in col_dict2.items(): 
            print(keys)
            # if()
            col_dict2[keys].append(tds[i].text)
            i=i+1
        # break
        
    print(col_dict2)
def getPuts(driver):
    main = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "optionChainTable-indices"))
    )
    tbod=main.find_element(By.TAG_NAME,"tbody")
    trs=tbod.find_elements(By.TAG_NAME,"tr")
    print(len(trs))
    for tr in trs:
        tds=tr.find_elements(By.TAG_NAME,"td")
        print(len(tds))
        # for td in tds:
        # if(len(tds)!=17):
        #     continue
        i=12
        for keys, values in puts_dict.items(): 
            # print(keys)
            # if()
            puts_dict[keys].append(tds[i].text)
            i=i+1
        # break 
        
    print(puts_dict)


col_dict={'SYMBOL':[],
'OPEN':[],
'HIGH':[],
'LOW':[],
'Prev. Close':[],
'LTP':[],
'change':[],
'%CHNG':[],
'Volume':[],
'Value':[],
'52W H':[],
'52W L':[],

}
col_dict2={
    "OI":[],
    "Chng in OI":[],
    "Volume":[],
    "IV":[],
    "LTP":[],
    "Chng":[],
    "Bid Qty":[],
    "Bid":[],
    "Ask":[],
    "Ask Qty":[],
    "Strike":[],
}
puts_dict={
    "Bid Qty":[],
    "Bid":[],
    "Ask":[],
    "Ask Qty":[],
    "Chng":[],
    "LTP":[],
    "IV":[],
    "Volume":[],
    "Chng in OI":[],
    "OI":[],
}
urls=["https://www.nseindia.com/market-data/live-equity-market?symbol=NIFTY%2050","https://www.nseindia.com/option-chain"]
if __name__ == '__main__':
    
    
    
    options = webdriver.FirefoxOptions()
    # options.add_argument("--headless")
    # options.add_argument('--disable-gpu')   
    start = int(sys.argv[1])
    # end = int(sys.argv[3])
    driver = webdriver.Firefox(options=options)

    
    # for i in range(start,end+1):

    
    driver.get(urls[start])
    if(start==0):
        getTable(driver)
        df=pd.DataFrame(col_dict)
        df.to_excel(f"data.xlsx")
            
        driver.quit()
    else:
        getOptions(driver)
        getPuts(driver)
    # time.sleep(20)
        df2=pd.DataFrame(puts_dict)
        df=pd.DataFrame(col_dict2)
        df2.to_excel(f"puts.xlsx")
        df.to_excel(f"calls.xlsx")
            
        driver.quit()
    # df=pd.DataFrame(col_dict)
    # df.index = np.arange(1, len(df) + 1)
    # df.to_excel(f"{name_dict[sys.argv[1]]}_data/data_{i}.xlsx")
    
    
    # path="unique_data"
    # isExist = os.path.exists(path)
    # if not isExist:
    #     os.makedirs(path)
    
    
