import os
import random
import time

import xlwt
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait

from front_login import *
from readConfig import ReadConfig
from db import DbOperate
from selenium.webdriver.chrome.options import Options
from mysqldb import connect

chrome_options = Options()
chrome_options.add_argument('--headless')
driver = webdriver.Chrome(chrome_options=chrome_options)

# driver = webdriver.Chrome()

driver.maximize_window()
driver.get(ReadConfig().get_root_url())
driver.get(ReadConfig().get_root_url())


class Common(object):
    def __init__(self):
        self.driver = driver
        # Excel写入
        self.row = 0
        self.workbook = xlwt.Workbook(encoding='utf-8')
        self.booksheet = self.workbook.add_sheet('Sheet1')
        self.timetemp = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())  # 存储Excel表格文件名编号
        # 每个案件的数量
        self.number = 1
        self.report_path = ReadConfig().save_report()
        self.windows = None
        self.screen_path = ReadConfig().save_screen()

    # 增加案件数量
    def number_add(self):
        if self.number > 1:
            for i in range(self.number):
                self.driver.find_element_by_xpath("//a[@class='add']").click()
        else:
            self.driver.find_element_by_xpath("//a[@class='add']").click()

    # 减少案件数量至1
    def number_minus(self):
        while self.number > 1:
            self.driver.find_element_by_xpath("//a[@class='jian']").click()

    # 存入数据库
    def save_to_mysql(self, parm):
        code = 0
        if isinstance(parm, list):
            parm.append(code)
        else:
            parm = list(parm)
            parm.append(code)
        res_code = connect(parm)
        print("存储状态", res_code)


    # 执行下单
    def execute_function(self, callback):
        try:
            eval("self.{}()".format(callback))
        except Exception as e:
            print("错误信息:", e)
            self.write_error_log(callback)
            time.sleep(0.5)
            self.write_error_log(str(e))

    def write_error_log(self, info):
        error_log_path = os.path.join(self.report_path,
                                      "error_log_{}.log".format(time.strftime("%Y-%m-%d", time.localtime())))
        with open(error_log_path, "a", encoding="utf-8") as f:
            f.write("{}: ".format(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())) + info + "\n")

    # 处理价格字符
    def process_price(self, price):
        if "￥" in price:
            price = price.replace("￥", '')
        return price

    # 关闭窗口
    def closed_windows(self, num):
        self.windows = self.driver.window_handles
        for n in range(num + 1, len(self.windows)):
            self.driver.switch_to.window(self.windows[n])
            self.driver.close()
        self.windows = self.driver.window_handles
        self.driver.switch_to.window(self.windows[num])

    # 存储信息
    def excel_number(self, infos):
        # 获取案件名称、案件号
        if infos:
            n = 0
            for info in infos:
                self.booksheet.write(self.row, n, info)
                self.booksheet.col(n).width = 300 * 28
                n += 1
            path = os.path.join(self.report_path, "report_{}.xls".format(self.timetemp))
            self.workbook.save(path)

    # 窗口截图
    def qr_shotscreen(self, windows_handle, name):
        current_window = self.driver.current_window_handle
        if current_window != windows_handle:
            self.driver.switch_to.window(windows_handle)
            path = self.screen_path
            self.driver.save_screenshot(path + self.timetemp + name + ".png")
            print("截图成功")
            self.driver.switch_to.window(current_window)
        else:
            path = self.screen_path
            self.driver.save_screenshot(path + self.timetemp +name + ".png")
            print("截图成功")
