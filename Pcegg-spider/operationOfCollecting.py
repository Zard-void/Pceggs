from operationOnPage import *
def getInfoOfCurPage(driver,wb):
    table = findObjSafely(driver, 'id', 'panel')
    trlist = table.find_elements_by_tag_name('tr')
    lenList = len(trlist)
    for index in range(lenList):
        listData = []
        table = findObjSafely(driver, 'id', 'panel')
        trlist = table.find_elements_by_tag_name('tr')
        # 遍历行对象，获取每一个行中所有的列对象
        tdlist = trlist[index].find_elements_by_tag_name('td')
        # 0:期号 1:时间 2:号码 3:人数 4:总额 5:中 6:我中 7:状态
        if (len(tdlist) < 8):
            break;          #已经超出表格，结束当前页，进入下一页
        if (tdlist[7].text == "已揭晓"):
            listData.append(tdlist[0].text)     #期号
            tagContent = tdlist[2].text
            listData.append(tagContent.split("+", 2)[0].split(" ", 1)[0])       #第一个数字
            listData.append(tagContent.split("+", 2)[1].split(" ", 1)[1])       #第二个数字
            listData.append(tagContent.split("+", 2)[2].split(" ", 2)[1])       #第三个数字
            enterResult(driver,tdlist[7])
            tagNum = findObjSafely(driver, 'xpath', '//td[contains(text(),"期号")]//font').text
            if (tagNum == listData[0]):     #结果界面期号与table界面期号一致
                tableMain = findObjSafely(driver, 'xpath', '//table[ @ cellspacing = "1"]')
                tableRowList = tableMain.find_elements_by_tag_name('tr')
                index = 0
                for childRow in tableRowList:
                    tableColumnList = childRow.find_elements_by_tag_name('td')
                    if (isObjExist(tableColumnList[0], 'tag_name', 'img')):
                        img = findObjSafely(tableColumnList[0], 'tag_name', 'img')
                        imgContent = img.get_attribute("src")
                        tmpStr = "number_" + str(index) + ".gif"
                        if tmpStr in imgContent:
                            listData.append(tableColumnList[1].text)
                            index = index + 1
            else:
                print("OMG")
            writeDataToExcel(driver,wb,listData)