from findObjects import *
import json
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

def loginWithCookie(driver):
    driver.delete_all_cookies()
    with open('cookies.txt', 'r', encoding='utf8') as f:
        listCookies = json.loads(f.read())
    for cookie in listCookies:
        driver.add_cookie(cookie)
    driver.refresh()
    return

def isLogin(driver):
    tickFor1S = 0
    while isObjExist(driver, 'id', 'fc_loginHead') == True and tickFor1S < 5:
        time.sleep(1)
        tickFor1S = tickFor1S + 1
    if isObjExist(driver,'id','fc_loginHead') == True:
        return False
    else:
        return True