import time
from selenium import webdriver
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
            boj = driver.find_element_by_tag_name(str)
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
        print(listCookies)
    for cookie in listCookies:
        driver.add_cookie(cookie)
    driver.refresh()
    return

def Test():
    wb = xw.Book("D:\\Test\\anjian.xlsm")
    sht0 = wb.sheets[0]
    value = sht0.range('A1').value
    print(value)
    value = sht0.range('A8').value
    print(value)


if __name__ == '__main__':
    if(0):
        driver = webdriver.Chrome()
        url = 'http://www.pceggs.com/'
        driver.get(url)
        if not isLogin(driver):
            loginWithCookie(driver)
        moveToStartPage(driver)
        table = findObjSafely(driver,'id','panel')
        trlist = table.find_elements_by_tag_name('tr')
        print(len(trlist))
        for row in trlist:
            #遍历行对象，获取每一个行中所有的列对象
            tdlist = row.find_elements_by_tag_name('td')
            for col in tdlist:
                print(col.text + '\t', end='')
            print('\n')
    else:
        Test()
