import time

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