import time
from selenium import webdriver


def login():
    acount_num = 'qq850224169'
    passwd_str = 'yzhch123456789'
    driver = webdriver.Chrome()
    url = 'http://www.pceggs.com/'
    driver.get(url)
    time.sleep(3)
    # 使用CSSSelector正则匹配头部
    #elem = driver.find_element_by_css_selector("iframe[id^='x-URS-iframe']")
    # 163登陆框是使用iframe进行嵌套的，所以需要先切换到该iframe
    #driver.switch_to.frame(elem)

    acount = driver.find_element_by_name('txt_UserName')
    acount.clear()
    acount.send_keys(acount_num)

    passwd = driver.find_element_by_name('txt_PWD')
    passwd.clear()
    passwd.send_keys(passwd_str)

    vercode = driver.find_element_by_name('txt_VerifyCode')
    vercode.clear()

    vercode_str = input("请输入验证码:\n")
    vercode.send_keys(vercode_str)

    time.sleep(3)
    click_button = driver.find_element_by_name('Login_Submit')
    click_button.click()
    time.sleep(5)
    cur_cookies = driver.get_cookies()[0]
    return cur_cookies


def isObjExist(driver, type, str):
    if type == 'name':
        obj = driver.find_element_by_name(str)
    elif type == 'id':
        obj = driver.find_element_by_id(str)

    if not obj :
        existance = True
    else:
        existance = False
    return existance


def Test():
    driver = webdriver.Chrome()
    url = 'http://www.pceggs.com/'
    driver.get(url)
    existance = isObjExist(driver, 'name', 'jajajajajajajjajaja')
    print(existance)

if __name__ == '__main__':
    Test()
