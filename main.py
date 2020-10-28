# This is a sample Python script.

# Press ⇧F10 to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import xlrd
import xlwt
from xlutils.copy import copy

def mergeExcelFile(aPath, aOutputPath, aSheetName):
    import os
    value_title = [
        ["公示编号", "姓名", "出生年月", "单位名称", "积分分值", "合法稳定就业", "合法稳定住所", "教育背景", "扣除取得学历（学位）期间累计的居住及就业分值", "职住区域", "创新创业",
         "纳税", "年龄", "荣誉表彰", "守法记录"], ]
    write_excel_xls(aOutputPath, aSheetName, value_title)

    for root,dirs,files in os.walk(aPath):
        total = 0
        for file in files:
            if "2020北京积分落户名单" not in file:
                continue

            #获取文件所属目录
            print(root)
            #获取文件路径
            print(os.path.join(root,file))
            filePath = os.path.join(root,file)
            data = xlrd.open_workbook(filePath)  # 文件名以及路径，如果路径或者文件名有中文给前面加一个r拜师原生字符。
            table = data.sheet_by_name(aSheetName)  # 通过名称获取

            rows = table.nrows
            total += rows - 1
            for i in range(1, rows):
                row_data = table.row_values(i)
                print(row_data)
                write_excel_xls_append(aOutputPath, [row_data])
        print("合并完毕，总行数为：" + str(total))


def write_excel_xls(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("xls格式表格写入数据成功！")


def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i + rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.

def start_get_data():
    from selenium import webdriver
    from selenium.webdriver.support.wait import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.by import By
    import time
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.support.select import Select  # select元素操作类

    import time

    driver = webdriver.Firefox()

    driver.get("http://fuwu.rsj.beijing.gov.cn/nwesqintegralpublic/settleperson/settlePersonTable")

    start_page = 1

    ## 获取 每页展示条数并设置为最大值
    element = driver.find_element_by_class_name("pageSize")
    select = Select(element)
    select.select_by_index(3)

    # driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")

    is_end = False

    total = 0

    print("开始处理数据")
    while not is_end:

        ## 获取当前激活页面数

        page_div = driver.find_element_by_class_name("pagination")
        current_page = page_div.find_element_by_tag_name("ul").find_element_by_class_name("active").find_element_by_tag_name("a").text


        if int(current_page) < start_page:
            page_info_list = page_div.find_element_by_tag_name(
                "ul").find_elements_by_tag_name("li")

            count = len(page_info_list)

            next_page = page_info_list[count - 2].find_element_by_tag_name("a")
            if next_page.text == ">":
                next_page.click()
                time.sleep(2)  # 打开网址后休息3秒钟,可用可不用
                print("当前页为 " + current_page + " 继续寻找下一页")
            continue

        book_name_xls = '2020北京积分落户名单_' + current_page +  '.xls'
        sheet_name_xls = '2020北京积分落户名单'
        value_title = [
            ["公示编号", "姓名", "出生年月", "单位名称", "积分分值", "合法稳定就业", "合法稳定住所", "教育背景", "扣除取得学历（学位）期间累计的居住及就业分值", "职住区域", "创新创业",
             "纳税", "年龄", "荣誉表彰", "守法记录"], ]
        write_excel_xls(book_name_xls, sheet_name_xls, value_title)

        ## 解析数据
        # 按行查询表格的数据，取出的数据是一整行，按空格分隔每一列的数据
        table_tr_body = driver.find_element_by_class_name("box").find_element_by_class_name("blue_table1").find_element_by_tag_name("tbody")
        rows = table_tr_body.find_elements_by_tag_name("tr")

        total += len(rows)

        for row in rows:
            # driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")
            value = []
            row_tds = row.find_elements_by_tag_name("td")
            value.append(row_tds[0].text)
            value.append(row_tds[1].text)
            value.append(row_tds[2].text)
            value.append(row_tds[3].text)
            value.append(row_tds[4].text)
            row_tds[5].click()
            time.sleep(1)
            handle = driver.current_window_handle  # 获得当前窗口,也就是弹出的窗口句柄,什么是句柄我也解释不清楚,反正它代表当前窗口
            driver.switch_to.window(handle)
            detail_model = driver.find_element_by_id("detailModal")
            detail_table_tr_body = detail_model.find_element_by_class_name(
                "blue_table1").find_element_by_tag_name("tbody")
            detail_rows = detail_table_tr_body.find_elements_by_tag_name("tr")
            for i in range(1, len(detail_rows)):
                row_tds = detail_rows[i].find_elements_by_tag_name("td")
                value.append(row_tds[2].text)

            try:
                detail_model.find_element_by_class_name("anniu").click()
            except ValueError:
                print("出错了")
                time.sleep(1)
                close_element = detail_model.find_element_by_class_name("m_close")
                driver.executeScript("arguments[0].click();", close_element)
                # detail_model.find_element_by_class_name("m_close").click()

            # driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")
            write_excel_xls_append(book_name_xls, [value])
            print(value)

        ## 完成一页
        page_info_list = driver.find_element_by_class_name("pagination").find_element_by_tag_name(
            "ul").find_elements_by_tag_name("li")

        count = len(page_info_list)

        next_page = page_info_list[count - 2].find_element_by_tag_name("a")
        if next_page.text == ">":
            next_page.click()
            time.sleep(2)  # 打开网址后休息3秒钟,可用可不用
        else:
            is_end = True
    print("结束，共处理" + str(total) + "条数据")



# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    ## 合并所有爬取的 excel 文件
    mergeExcelFile("/Users/lihui/Desktop/excel", r"/Users/lihui/Desktop/2020北京积分落户名单.xls", r"2020北京积分落户名单")

    ## 开始爬取数据
    # start_get_data()
