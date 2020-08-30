from findObjects import *
def moveToStartPage(driver):
    driver.get("http://www.pceggs.com/play/playIndex.aspx")
    driver.get("http://www.pceggs.com/play/pxya.aspx")
    return

def moveToPage(driver,page):
    form = findObjSafely(driver, 'id', 'pagination')
    pageInput = findObjSafely(driver, 'id', 'CurrentPageIndex')
    pageTag = "arguments[0].value = " + "'" + str(page) + "';"
    driver.execute_script(pageTag, pageInput)
    form.submit()

def enterResult(driver,tdlistMem):
    tagContent = findObjSafely(tdlistMem, 'tag_name', 'a')
    url = tagContent.get_attribute('href')
    driver.get(url)

def writeDataToExcel(driver,wb,listData):
    columnOfCell = 1
    rowOfCell = 2
    for obj in listData:
        wb.sheets['sheet1'].range(rowOfCell, columnOfCell).value = obj
        columnOfCell = columnOfCell + 1
        if (columnOfCell == 5):
            columnOfCell = 7
    wb.sheets['sheet1'].api.Rows(2).Insert()
    driver.back()