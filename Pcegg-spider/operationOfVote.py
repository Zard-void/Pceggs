from operationOnPage import *
import selenium

def findTheFirstVoted(driver,curlist,nextlist,nowRow):
    if (curlist[7].text == "猜猜" and nextlist[7].text == "已揭晓"):
        return nowRow
    elif(curlist[7].text == "猜猜"and nextlist[7].text == "猜猜"):
        return 0
    elif(curlist[7].text == "揭晓中"and nextlist[7].text == "已揭晓"):
        return -1
    elif(curlist[7].text == "已揭晓" and nextlist[7].text == "已揭晓"):
        raise Exception("未找到")
    else:
        return 0

def findRefreshButton(driver):
    button = driver.find_elements_by_xpath('//form//div[@class="touzhu1"]')
    for refreshButton in button:
        if(refreshButton.text == "刷新赔率"):
            return refreshButton


def enterVotePage(driver):
    driver.execute_script("window.scrollBy(0,1000)")
    table = findObjSafely(driver, 'id', 'panel')
    trlist = table.find_elements_by_tag_name('tr')
    lenList = len(trlist)
    for index in range(lenList):
        table = findObjSafely(driver, 'id', 'panel')
        trlist = table.find_elements_by_tag_name('tr')
        # 遍历行对象，获取每一个行中所有的列对象
        tdlist = trlist[index].find_elements_by_tag_name('td')
        if(index < 20):
            nextlist = trlist[index + 1].find_elements_by_tag_name('td')
            if(len(nextlist) == 8 and nextlist[7].text in "已揭晓，揭晓中，猜猜"):
                nowStatus = findTheFirstVoted(driver,tdlist, nextlist,index)
                if(nowStatus == -1):    #-1代表当前不能投注
                    return -1
                elif(nowStatus == 0):    #0代表当前行不是目标行
                    continue
                elif(nowStatus < 10):
                    tdlist[7].click()
                    #enterResult(driver,tdlist[7])
                    return 1

def getCurrentTime(driver):
    nowTimeInfo = driver.find_elements_by_xpath('//span//span')
    length = len(nowTimeInfo)
    for index in range(length):
        nowTimeInfo = driver.find_elements_by_xpath('//span//span')
        nowTime = nowTimeInfo[index]
        try:
            if(float(nowTime.text) < 3000):
                return nowTime.text
        except selenium.common.exceptions.StaleElementReferenceException:
            return -1

def getCurrentOdd(driver):
    currentOdd = []
    leftTable = findObjSafely(driver,'xpath',"// form // table[ @ width = '467'][ @ align = 'left']")
    trlist = leftTable.find_elements_by_tag_name("tr")
    trlen = len(trlist)
    for index in range(trlen):
        if index == 0:
            continue
        tdlist = trlist[index].find_elements_by_tag_name("td")
        currentOdd.append(float(tdlist[2].text))
    rightTable = findObjSafely(driver,'xpath',"// form // table[ @ width = '467'][ @ align = 'right']")
    trlist = rightTable.find_elements_by_tag_name("tr")
    trlen = len(trlist)
    for index in range(trlen):
        if index == trlen - 1:
            continue
        tdlist = trlist[trlen -1 -index].find_elements_by_tag_name("td")
        currentOdd.append(float(tdlist[2].text))
    for ls in currentOdd:
        if(ls == 0):
            return None
    return currentOdd

def fillOutTheForm(driver):
    normallist = [1000,333.3333,166.6667,100,66.66667,47.61905,35.71429,27.77778,22.22222,18.18182,15.87302,14.49275,13.69863,13.33333,
                  13.33333,13.69863,14.49275,15.87302,18.18182,22.22222,27.77778,35.71429,47.61905,66.66667,100,166.6667,333.3333,1000]

    leftTable = findObjSafely(driver, 'xpath', "// form // table[ @ width = '467'][ @ align = 'left']")
    trlist = leftTable.find_elements_by_tag_name("tr")
    trlen = len(trlist)
    for index in range(trlen):
        if index == 0:
            continue
        tdlist = trlist[index].find_elements_by_tag_name("td")
        if(float(tdlist[2].text) > normallist[index - 1]):
            try:
                tdlist[3].click()
            except selenium.common.exceptions.ElementClickInterceptedException:
                tdlist = trlist[index].find_elements_by_tag_name("td")
                tdlist[3].click()

    rightTable = findObjSafely(driver, 'xpath', "// form // table[ @ width = '467'][ @ align = 'right']")
    trlist = rightTable.find_elements_by_tag_name("tr")
    trlen = len(trlist)
    for index in range(trlen):
        if index == trlen - 1:
            continue
        tdlist = trlist[trlen - 1 - index].find_elements_by_tag_name("td")
        if (float(tdlist[2].text) > normallist[index + 14]):
            try:
                tdlist[3].click()
            except selenium.common.exceptions.ElementClickInterceptedException:
                tdlist = trlist[trlen - 1 - index].find_elements_by_tag_name("td")
                tdlist[3].click()
    return

def clickEnsureButton(driver):
    ensureButton = findObjSafely(driver, 'xpath', '//div[@class = "conform_btn"]')
    ensureButton.click()
    ensureButton = findObjSafely(driver,'id','fc_an_l170223')
    ensureButton.click()
    ensureButton = findObjSafely(driver,'id','fc_an_m170223')
    ensureButton.click()

def vote(driver):
    if(enterVotePage(driver) != -1):
        refreshButton = findRefreshButton(driver)
        refreshFlag = 0
        while(1):
            time.sleep(2)
            currentTime = int(getCurrentTime(driver))
            if(refreshFlag == 1):
                try:
                    refreshButton.click()
                except selenium.common.exceptions.StaleElementReferenceException:
                    refreshButton = findRefreshButton(driver)
            if(currentTime == -1):
                continue
            elif(8<currentTime < 30):
                if (refreshFlag == 0):
                    driver.refresh()
                refreshFlag = 1
                #currentOdd = getCurrentOdd(driver)
            elif(0<currentTime<8):
                fillOutTheForm(driver)
                clickEnsureButton(driver)
                return
