from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import xlwings as xw
from login import *
from operationOfCollecting import *
from operationOfVote import *


#chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenum\AutomationProfile"

def Test():
    chrome_options = Options()
    wb = xw.Book("D:\\Test\\database.xlsm")
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    chrome_driver = "D:\\CodingTools\\Python3.7.0\\chromedriver.exe"
    driver = webdriver.Chrome(options=chrome_options)
    url = 'http://www.pceggs.com/'
    driver.get(url)
    if not isLogin(driver):
        loginWithCookie(driver)
    moveToStartPage(driver)
    while(1):
        enterVotePage(driver)
        vote(driver)
        time.sleep(150)
        driver.refresh()


if __name__ == '__main__':
    if(1):
        listData = []
        chrome_options = Options()
        wb = xw.Book("D:\\Test\\database.xlsm")
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        chrome_driver  = "D:\\CodingTools\\Python3.7.0\\chromedriver.exe"
        driver = webdriver.Chrome(options=chrome_options)
        #driver = webdriver.Chrome()
        url = 'http://www.pceggs.com/'
        driver.get(url)
        if not isLogin(driver):
            login(driver)
        moveToStartPage(driver)
        theEndNum = 1037206
        nowPage = 1
        while(nowPage < 25):
            getInfoOfCurPage(driver,wb)
            nowPage = nowPage + 1
            moveToPage(driver, nowPage)
    else:
        while(1):
            try:
                Test()
            except selenium.common.exceptions.ElementNotInteractableException:
                continue
            except ValueError:
                continue

