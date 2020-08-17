import time
from selenium import webdriver


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

    click_button = findObjSafely(driver,'name','Login_Submit')
    click_button.click()
    time.sleep(5)
    cur_cookies = driver.get_cookies()[0]
    return cur_cookies

def findObjSafely(driver,type,str):
    waitForObjExist(driver,type,str)
    if(type == 'name'):
        return driver.find_element_by_name(str)
    elif(type == 'id'):
        return driver.find_element_by_id(str)



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
def Test():
    driver = webdriver.Chrome()
    url = 'http://www.pceggs.com/'
    driver.get(url)
    existance = isObjExist(driver, 'name', 'txt_UserName')
    print(existance)

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
            handleObjExistErr()

def handleObjExistErr():
    return


if __name__ == '__main__':
    driver = webdriver.Chrome()
    url = 'http://www.pceggs.com/'
    driver.get(url)
    if not isLogin(driver):
        login(driver)

