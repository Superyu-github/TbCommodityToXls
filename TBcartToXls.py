from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException  # 导入NoSuchElementException
import xlrd
import xlwt


# 获取商品详情页数据
# 参数select_switch为选择导出功能的开关
# 由于只有待收货状态的商品才支持选择，为方便统计已收获商品，故设置此功能开关
# select_switch为False时导出当前页所有商品
def get_order_data(driver, select_switch=True):
    tittle = []
    item = []
    price = []
    amount = []
    link = []
    totle_price = []
    for i in range(4, 18 + 1):
        for j in range(1, 50):
            try:
                is_selected = driver.find_element_by_xpath(
                    f'//*[@id="tp-bought-root"]/div[{i}]/div/table/tbody[1]/tr/td[1]/label/span[1]/input').is_selected()
                if is_selected or (not select_switch):
                    # 获取标题和链接
                    tittle_link = driver.find_element_by_xpath(
                        f'//*[@id="tp-bought-root"]/div[{i}]/div/table/tbody[2]/tr[{j}]/td[1]/div/div[2]/p[1]/a[1]')
                    tittle.append(tittle_link.get_attribute('text'))
                    link.append(tittle_link.get_attribute('href'))
                    # 获取总价
                    try:
                        totle_price.append(driver.find_element_by_xpath(
                            f'//*[@id="tp-bought-root"]/div[{i}]/div/table/tbody[2]/tr[1]/td[5]/div/div[1]/p/strong/span[2]').text)
                    except NoSuchElementException:
                        totle_price.append("")
                    # 获取单价
                    try:
                        price.append(driver.find_element_by_xpath(
                            f'//*[@id="tp-bought-root"]/div[{i}]/div/table/tbody[2]/tr[{j}]/td[2]/div/p/span[2]').text)
                    except NoSuchElementException:
                        price.append("")
                    # 获取数量
                    try:
                        amount.append(driver.find_element_by_xpath(
                            f'//*[@id="tp-bought-root"]/div[{i}]/div/table/tbody[2]/tr[{j}]/td[3]/div/p').text)
                    except NoSuchElementException:
                        amount.append("")
                    # 获取商品详情
                    try:
                        item.append(driver.find_element_by_xpath(
                            f'//*[@id="tp-bought-root"]/div[{i}]/div/table/tbody[2]/tr[{j}]/td[1]/div/div[2]/p[2]/span/span[3]').text)
                    except NoSuchElementException:
                        item.append("")
            except NoSuchElementException:
                break
    return tittle, item, price, amount, link, totle_price


# 获取购物车数据
def get_cart_data(driver):
    title = []
    item = []
    price = []
    amount = []
    link = []
    for i in range(1, 30):
        for j in range(1, 50):
            tianmao = f"/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[{i}]/div/div[2]/div/div/div[2]/div/div[{j}]/div/ul/"  # 天猫店铺索引方式
            taobao = f'/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[{i}]/div/div[2]/div/div/div[{j}]/div/ul/'  # 淘宝店铺索引方式
            try:
                # 判断是否选中
                try:
                    is_selected = driver.find_element_by_xpath(tianmao + f'li[1]/div/div/div/input').is_selected()
                except NoSuchElementException:
                    is_selected = driver.find_element_by_xpath(taobao + f'li[1]/div/div/div/input').is_selected()
                if is_selected:
                    # 获取标题和链接
                    try:
                        title_link = driver.find_element_by_xpath(tianmao + f'li[2]/div/div[2]/div[1]/a')
                    except NoSuchElementException:
                        title_link = driver.find_element_by_xpath(
                            taobao + f'li[2]/div/div[2]/div[1]/a')  # 若天猫索引方式报错，尝试淘宝索引方式
                    title.append(title_link.get_attribute('text'))  # 获取商品标题
                    link.append(title_link.get_attribute('href'))  # 获取商品链接
                    # 获取单价
                    try:
                        try:
                            price.append(driver.find_element_by_xpath(tianmao + f'li[4]/div/div/div/div/em').text)
                        except NoSuchElementException:
                            price.append(driver.find_element_by_xpath(
                                taobao + f'li[4]/div/div/div/div/em').text)  # 若天猫索引方式报错，尝试淘宝索引方式
                    except NoSuchElementException:
                        price.append("")  # 两种方式尝试均报错，说明没有此元素，留空
                    # 获取详情
                    try:
                        try:
                            item.append(driver.find_element_by_xpath(tianmao + f'li[3]/div/p').text)
                        except NoSuchElementException:
                            item.append(driver.find_element_by_xpath(taobao + f'li[3]/div/p').text)
                    except NoSuchElementException:
                        item.append("")
                    # 获取数量
                    try:
                        try:
                            amount.append(
                                driver.find_element_by_xpath(tianmao + f'li[5]/div/div/div[1]/input').get_attribute(
                                    "value"))
                        except NoSuchElementException:
                            amount.append(
                                driver.find_element_by_xpath(taobao + f'li[5]/div/div/div[1]/input').get_attribute(
                                    "value"))
                    except NoSuchElementException:
                        amount.append("")

            except NoSuchElementException:
                break
    return title, item, price, amount, link


# 创建工作表
def craet_workbook(type, path="data.xls"):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Sheet1")
    if type == '1':
        sheet.write(0, 0, "商品名")
        sheet.write(0, 1, "商品详情")
        sheet.write(0, 2, "单价")
        sheet.write(0, 3, "数量")
        sheet.write(0, 4, "链接")
    else:
        sheet.write(0, 0, "商品名")
        sheet.write(0, 1, "商品详情")
        sheet.write(0, 2, "单价")
        sheet.write(0, 3, "数量")
        sheet.write(0, 4, "店铺实付款")
        sheet.write(0, 5, "链接")
    workbook.save(path)
    return workbook


if __name__ == "__main__":

    select = ''
    switch = ''
    login = ''
    filter = ''
    type = input("选择需要抓取数据的页面(1:购物车 2:商品详情):")
    driver = webdriver.Chrome("./chromedriver.exe")
    if type == '1':
        driver.get('https://cart.taobao.com/cart.htm')
    else:
        driver.get('https://buyertrade.taobao.com/trade/itemlist/list_bought_items.htm')
    workbook = craet_workbook(type)  # 创建工作表
    sheet = workbook.get_sheet("Sheet1")

    while login != "y":
        login = input("登录完成？(y/n)")

    if type == '2':
        print("商品详情页仅支持未收货产品选择导出，若要导出已收货产品，请将选择开关关闭，进行全局页面导出")
        switch = input("打开选择开关？(1开 0关)")

    while select != "y":
        select = input("请选择需要导出的商品，选择完成？(y/n)")
    count = 1
    filter = input("过滤0元订单？(y/n)")
    print("正在导出，请稍后......")
    if type == "1":
        title, item, price, amount, link = get_cart_data(driver)
        for i in range(len(title)):
            if filter == 'y':
                if price[i] == "0.00":
                    continue
            sheet.write(count, 0, title[i])
            sheet.write(count, 1, item[i])
            sheet.write(count, 2, price[i])
            sheet.write(count, 3, amount[i])
            sheet.write(count, 4, link[i])
            count += 1
    else:
        title, item, price, amount, link, totle_price = get_order_data(driver, bool(switch))
        for i in range(len(title)):
            if filter == 'y':
                if price[i] == "0.00":
                    continue
            sheet.write(count, 0, title[i])
            sheet.write(count, 1, item[i])
            sheet.write(count, 2, price[i])
            sheet.write(count, 3, amount[i])
            sheet.write(count, 4, totle_price[i])
            sheet.write(count, 5, link[i])
            count += 1
    workbook.save("data.xls")
    print("导出完毕")




