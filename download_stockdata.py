
### import some required packages

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
import re
from selenium.webdriver.common.action_chains import ActionChains


### Connecting to website

driver_path = ChromeDriverManager().install()
print(driver_path)
service = Service(driver_path)
driver = webdriver.Chrome(service=service)
driver.get("https://www.twse.com.tw/zh/trading/historical/mi-index.html")
print(driver.title)


### Set up buttons,download all data

select_yy= driver.find_element(By.NAME, 'yy')
select_mm= driver.find_element(By.NAME, 'mm')
select_dd= driver.find_element(By.NAME, 'dd')
select_type= driver.find_element(By.NAME, 'type')
Select(select_type).select_by_index(3)
clickbutton_result = driver.find_element(By.XPATH, '//*[@id="form"]/div/div[1]/div[3]/button')
clickbutton_dlcsv = driver.find_element(By.XPATH, '//*[@id="reports"]/div[1]/button[2]')


### set month

big_month = [0, 2, 4, 6, 7, 9, 11]
print(big_month)


### Write a loop

for yys in range(11, 21):
    Select(select_yy).select_by_index(yys)
    for mms in big_month:
        #這邊有改月份，注意
        Select(select_mm).select_by_index(mms)
        #如果出問題，從這邊開始改
        for dds in range(0, 31):
            Select(select_dd).select_by_index(dds)
            clickbutton_result.click()
            time.sleep(10)
            try:
                clickbutton_dlcsv.click()
            except:
                print("沒有資料")
                continue








