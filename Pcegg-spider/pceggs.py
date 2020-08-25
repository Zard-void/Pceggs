import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import json
import xlwings as xw

def login(driver):
    acount_num = 'qq850224169'
    passwd_str = 'yzhch123456789'
    # 使用CSSSelector正则匹配头部
    #elem = driver.find_element_by_css_selector("iframe[id^='x-URS-iframe']")
    # 163登陆框是使用iframe进行嵌套的，所以需要先切换到该iframe
    #driver.switch_to.frame(elem)

    acount = findObjSafely(driver,'name','txt_UserName')
    acount.clear()
    acount.send_keys(acount_num)

    passwd = findObjSafely(driver,'name','txt_PWD')
    passwd.clear()
    passwd.send_keys(passwd_str)

    vercode = findObjSafely(driver,'name','txt_VerifyCode')
    vercode.clear()
    vercode_str = input("请输入验证码:\n")
    vercode.send_keys(vercode_str)

    loginWay = findObjSafely(driver,'id','LoginWay')
    loginWay.click()

    click_button = findObjSafely(driver,'name','Login_Submit')
    click_button.click()
    time.sleep(5)
    curCookies = driver.get_cookies()
    jsonCookies = json.dumps(curCookies)

    with open('cookies.txt', 'w') as cookief:
        cookief.seek(0)
        cookief.truncate()
        cookief.write(jsonCookies)
    return curCookies

def findObjSafely(driver,type,str):
    if(waitForObjExist(driver,type,str)):
        if(type == 'name'):
            return driver.find_element_by_name(str)
        elif(type == 'id'):
            return driver.find_element_by_id(str)
        elif(type == 'tag_name'):
            return driver.find_element_by_tag_name(str)
        elif(type == 'xpath'):
            return driver.find_element_by_xpath(str)
    else:
        return None



def isObjExist(driver, type, str):
    from selenium.common.exceptions import NoSuchElementException
    if type == 'name':
        try:
            obj = driver.find_element_by_name(str)
        except NoSuchElementException:
            return False
        else:
            return True
    elif type == 'id':
        try:
            obj = driver.find_element_by_id(str)
        except NoSuchElementException:
            return False
        else:
            return True
    elif type == 'tag_name':
        try:
            obj = driver.find_element_by_tag_name(str)
        except NoSuchElementException:
            return False
        else:
            return True
    elif type == 'xpath':
        try:
            obj = driver.find_element_by_xpath(str)
        except NoSuchElementException:
            return False
        else:
            return True



def isLogin(driver):
    tickFor1S = 0
    while isObjExist(driver,'id','fc_loginHead') == True and tickFor1S < 5:
        time.sleep(1)
        tickFor1S = tickFor1S + 1
    if isObjExist(driver,'id','fc_loginHead') == True:
        return False
    else:
        return True

def waitForObjExist(driver,type,str):
    tickFor1S = 0
    waitTime = 5
    while isObjExist(driver,type,str) == False and tickFor1S < waitTime:
        time.sleep(1)
        tickFor1S = tickFor1S + 1
    try:
        if tickFor1S >= waitTime:
            raise Exception(type + '不存在')
    except:
        return False
    return True


def moveToStartPage(driver):
    driver.get("http://www.pceggs.com/play/playIndex.aspx")
    driver.get("http://www.pceggs.com/play/pxya.aspx")
    return

def loginWithCookie(driver):
    driver.delete_all_cookies()
    with open('cookies.txt', 'r', encoding='utf8') as f:
        listCookies = json.loads(f.read())
    for cookie in listCookies:
        driver.add_cookie(cookie)
    driver.refresh()
    return

def Test():
    wb = xw.Book("D:\\Test\\database.xlsm")
    wb.sheets['sheet1'].range('A2').value='人生'
    wb.sheets['sheet1'].api.Rows(2).Insert()



if __name__ == '__main__':
    if(1):
        listData = []
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        chrome_driver  = "D:\\CodingTools\\Python3.7.0\\chromedriver.exe"
        driver = webdriver.Chrome(options=chrome_options)
        url = 'http://www.pceggs.com/'
        driver.get(url)
        if not isLogin(driver):
            loginWithCookie(driver)
        moveToStartPage(driver)
        table = findObjSafely(driver,'id','panel')
        trlist = table.find_elements_by_tag_name('tr')
        for row in trlist:
            #遍历行对象，获取每一个行中所有的列对象
            tdlist = row.find_elements_by_tag_name('td')
            #0:期号 1:时间 2:号码 3:人数 4:总额 5:中 6:我中 7:状态
            if (tdlist[7].text == "已揭晓"):
                listData.append(tdlist[0].text)
                tagContent = tdlist[2].text
                listData.append(tagContent.split("+", 2)[0].split(" ",1)[0])
                #print(tagContent.split("+",2)[0].split(" ",1)[0])
                listData.append(tagContent.split("+", 2)[1].split(" ",1)[1])
                #print(tagContent.split("+", 2)[1].split(" ",1)[1])
                listData.append(tagContent.split("+", 2)[2].split(" ",2)[1])
                #print(tagContent.split("+", 2)[2].split(" ",2)[1])
                #print("已揭晓")
                tdlist[7].click()
                tagContent = findObjSafely(driver,'xpath','//td[contains(text(),"期号")]')
                tagNum = tagContent.text.split("：",1)[1]
                if (tagNum == listData[0]):
                    tableMain = findObjSafely(driver,'xpath','//table[ @ cellspacing = "1"]')
                    tableRowList = tableMain.find_elements_by_tag_name('tr')
                    index = 0
                    for childRow in tableRowList:
                        tableColumnList = childRow.find_elements_by_tag_name('td')
                        if(isObjExist(tableColumnList[0],'tag_name','img')):
                            img = findObjSafely(tableColumnList[0],'tag_name','img')
                            imgContent = img.get_attribute("src")
                            tmpStr = "number_" + str(index) + ".gif"
                            if tmpStr in imgContent:
                                listData.append(tableColumnList[1].text)
                                index = index + 1
                else:
                    print("OMG")
                for obj in listData:
                    print(obj)
                time.sleep(30)
    else:
        Test()
