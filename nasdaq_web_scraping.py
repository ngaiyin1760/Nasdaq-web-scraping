####################################################
################# NASDAQ DATA LINK #################
####################################################

# import necessary library
#import nasdaqdatalink

# set start and end date for data request
#start = '2015-01-01'
#end = '2021-12-31'

# send request to extract time series data
#data = nasdaqdatalink.get('WIKI/AAPL',
#                           start_date = start,
#                           end_date = end)


####################################################
########## WEB SCRAPPING TO DOWNLOAD DATA ##########
####################################################

# import necessary library
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# open the Edge driver
driver = webdriver.Edge("C:/Users/User/Desktop/edgedriver_win32/msedgedriver.exe")

# set nasdaq url
nasdaq_url = 'https://www.nasdaq.com/market-activity/stocks/'

# set tickers list
ticker_list = ['AAPL', 'AMZN', 'GOOGL', 'MSFT', 'TSLA']

for ticker in ticker_list:
    
    link = nasdaq_url + ticker + "/historical"
    
    driver.get(link) # control the edge browser to navigate the corresponding 
    driver.maximize_window() # maximize the window to prevent clicking wrong button
    
    ########## Click "Accept all cookies" to proceed ##########
    if ticker_list.index(ticker) == 0: # only need to click the "Accept all cookies" once when the browser is newly opened
        wait = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@id='onetrust-button-group']//button[@id='onetrust-accept-btn-handler']")))   
        cookies_button = driver.find_element(By.XPATH, "//div[@id='onetrust-button-group']//button[@id='onetrust-accept-btn-handler']")
        cookies_button.click()
        
    
    ########## Click "MAX" button ##########
    wait = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='table-tabs__list']//button[@class='table-tabs__tab' and text() = 'MAX']")))   
    max_button = driver.find_element(By.XPATH, "//div[@class='table-tabs__list']//button[@class='table-tabs__tab' and text() = 'MAX']")
    max_button.click()
    
    
    ########## Click "Download Data" button ##########
    wait = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='historical-data__controls']//button[@class='historical-data__controls-button--download historical-download']//*[name()='svg']")))   
    download_button = driver.find_element(By.XPATH, "//div[@class='historical-data__controls']//button[@class='historical-data__controls-button--download historical-download']//*[name()='svg']")
    download_button.click()
    
    print('Finished downloading historical price data of ' + ticker)

driver.close() # close the Edge browser
time.sleep(15) # wait until the last excel file downloaded
      

##############################################
######### COMBINE ALL EXCEL'S DATA ###########
##############################################

# import necessary library
import os
import pandas as pd

# set the path that store the downloaded excel
download_path = 'C:/Users/User/Downloads/'

# set a dictionary to store all stock data
data_path = {}
i = 0

# set up an empty excel workbook to consolidate all the results
wb_final = pd.ExcelWriter("price_data.xlsx", engine='xlsxwriter')

for file in os.listdir(download_path):
    if file.split('_')[0] == 'HistoricalData':
        
        excel_path = download_path + file
        df = pd.read_csv(excel_path, index_col='Date')
        df.to_excel(wb_final, sheet_name=ticker_list[i])        
        i += 1

# save the consolidated workbook
wb_final.save()
print('Finished consolidating the price data for all tickers')









