import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

def open_site(driver):
    driver.get('https://in.tradingview.com/screener/') 

def maximize(driver):
    driver.maximize_window()

def getstocks(driver):
    alldata = driver.find_element_by_name("tv-screener__symbol")
    print(alldata)

if __name__=="__main__":
    webdriverpath="D:\software\chrome webdriver\chromedriver.exe"
    driver = webdriver.Chrome(webdriverpath)# Optional argument, if not specified will search path.
    open_site(driver)
    # getstocks(driver)