import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from db import DbOperate
from Common import Common
from mysqldb import connect
from selenium import webdriver
from readConfig import ReadConfig
from selenium.webdriver.chrome.options import Options
import os

class FunctionName(type):
    def __new__(cls, name, bases, attrs, *args, **kwargs):
        count = 0
        attrs["__Func__"] = []
        for k, v in attrs.items():
            if "copyright_" in k:
                attrs["__Func__"].append(k)
                count += 1

        attrs["__FuncCount__"] = count
        return type.__new__(cls, name, bases, attrs)


chrome_options = Options()
chrome_options.add_argument('--headless')
driver = webdriver.Chrome(chrome_options=chrome_options)

# driver = webdriver.Chrome()

driver.maximize_window()
driver.get(ReadConfig().get_root_url())
driver.get(ReadConfig().get_root_url())


class Execute(object, metaclass=FunctionName):
    def __init__(self):
        self.driver = driver
        self.timetemp = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())  # 存储Excel表格文件名编号
        # 每个案件的数量
        self.number = 1
        self.dboperate = DbOperate()
        self.db = "copyright"
        self.windows = None
        self.report_path = ReadConfig().save_report()

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

    def write_error_log(self, info):
        error_log_path = os.path.join(self.report_path,
                                      "error_log_{}.log".format(time.strftime("%Y-%m-%d", time.localtime())))
        with open(error_log_path, "a", encoding="utf-8") as f:
            f.write("{}: ".format(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())) + info + "\n")

    # 执行下单
    def execute_function(self, callback):
        try:
            eval("self.{}()".format(callback))
        except Exception as e:
            print("错误信息:", e)
            self.write_error_log(callback)
            time.sleep(0.5)
            self.write_error_log(str(e))

    # 
    def closed_windows(self, num):
        self.windows = self.driver.window_handles
        for n in range(num + 1, len(self.windows)):
            self.driver.switch_to.window(self.windows[n])
            self.driver.close()
        self.windows = self.driver.window_handles
        self.driver.switch_to.window(self.windows[num])

    # 计算机软件著作权登记
    def copyright_computer_software_01(self):
        all_type = [u'计算机软件著作权登记']
        type_code = ["computer"]
        for index, copyright_type in enumerate(all_type):
            if self.dboperate.exists(type_code[index]):
                try:
                    locator = (By.XPATH, "//div[@class='isnav-first']/div[1]/h2")
                    WebDriverWait(self.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
                    aa = self.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[3]")
                    ActionChains(self.driver).move_to_element(aa).perform()
                    self.driver.find_element_by_link_text(copyright_type).click()
                    # 切换至新窗口
                    self.windows = self.driver.window_handles
                    self.driver.switch_to.window(self.windows[-1])
                    # 服务类型：
                    for num in range(1, 7):
                        if self.dboperate.is_member(type_code[index], num):
                            self.driver.find_element_by_xpath("//ul[@p='232']/li[{}]/a".format(num)).click()
                            case_name = self.driver.find_element_by_xpath(
                                "//ul[@p='232']/li[{}]/a".format(num)).text
                            case_name = "-".join([str(copyright_type), case_name])

                            # 数量加减
                            # self.common.number_add()
                            # self.common.number_minus()
                            time.sleep(0.5)
                            while not self.driver.find_element_by_id("totalfee").is_displayed():
                                time.sleep(0.5)
                            # 获取详情页 价格
                            detail_price = self.driver.find_element_by_xpath("(.//div[@class='sames']//label[@id='totalfee'])").text
                            print("{}详情页价格".format(case_name), detail_price)
                            self.dboperate.del_elem(type_code[index], num)
                            self.save_to_mysql([case_name, detail_price])
                    self.closed_windows(0)
                except Exception as e:
                    print(e)
                    self.driver.switch_to.window(self.windows[0])
        time.sleep(1)

    # 美术作品著作权登记-30日
    def copyright_art_works_01(self):
        all_type = [u'美术作品著作权登记']
        type_code = ["art"]
        for index, copyright_type in enumerate(all_type):
            if self.dboperate.exists(type_code[index]):
                try:
                    locator = (By.XPATH, "//div[@class='isnav-first']/div[1]/h2")
                    WebDriverWait(self.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
                    aa = self.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[3]")
                    ActionChains(self.driver).move_to_element(aa).perform()
                    self.driver.find_element_by_link_text(copyright_type).click()
                    # 切换至新窗口
                    self.windows = self.driver.window_handles
                    self.driver.switch_to.window(self.windows[-1])
                    for num in range(1, 7):
                        if self.dboperate.is_member(type_code[index], num):
                            self.driver.find_element_by_xpath("//ul[@p='107538']/li[{}]/a".format(num)).click()
                            case_name = self.driver.find_element_by_xpath(
                                "//ul[@p='107538']/li[{}]/a".format(num)).text
                            case_name = "-".join([str(copyright_type), case_name])
                            
                            # 数量加减
                            # self.common.number_add()
                            # # self.common.number_minus()
                            time.sleep(0.5)
                            while not self.driver.find_element_by_id("totalfee").is_displayed():
                                time.sleep(0.5)
                            # 获取详情页 价格
                            detail_price = self.driver.find_element_by_xpath("(.//div[@class='sames']//label[@id='totalfee'])").text
                            print("{}详情页价格".format(case_name), detail_price)

                            self.dboperate.del_elem(type_code[index], num)
                            self.save_to_mysql([case_name, detail_price])
                    self.closed_windows(0)

                except Exception as e:
                    print(e)
                    self.driver.switch_to.window(self.windows[0])
        time.sleep(1)

    # 文字作品著作权登记
    def copyright_writings_01(self):
        # 选择文字作品著作权登记
        all_type = [u'汇编作品著作权登记', u'文字作品著作权登记', u'摄影作品著作权登记', u'电影作品著作权登记', u'音乐作品著作权登记', u'曲艺作品著作权登记']
        type_code = ["compile", "word", "photography", "film", "music", "drama"]
        for index, copyright_type in enumerate(all_type):
            if self.dboperate.exists(type_code[index]):
                try:
                    locator = (By.XPATH, "//div[@class='isnav-first']/div[1]/h2")
                    WebDriverWait(self.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
                    aa = self.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[3]")
                    ActionChains(self.driver).move_to_element(aa).perform()
                    self.driver.find_element_by_link_text(copyright_type).click()
                    # 切换至新窗口
                    self.windows = self.driver.window_handles
                    self.driver.switch_to.window(self.windows[-1])
                    # 案件类型：
                    for num in range(1, 7):
                        if self.dboperate.is_member(type_code[index], num):
                            self.driver.find_element_by_xpath("//ul[@id='ulType']/li[{}]/a".format(num)).click()
                            case_name = self.driver.find_element_by_xpath(
                                "//ul[@id='ulType']/li[{}]/a".format(num)).text
                            case_name = "-".join([str(copyright_type), case_name])

                            # 数量加减
                            # self.common.number_add()
                            # # self.common.number_minus()
                            time.sleep(0.5)
                            while not self.driver.find_element_by_id("totalfee").is_displayed():
                                time.sleep(0.5)
                            # 获取详情页 价格
                            detail_price = self.driver.find_element_by_xpath("(.//div[@class='sames']//label[@id='totalfee'])").text
                            print("{}详情页价格".format(case_name), detail_price)
                            self.dboperate.del_elem(type_code[index], num)
                            self.save_to_mysql([case_name, detail_price])
                    self.closed_windows(0)

                except Exception as e:
                    print(e)
                    self.driver.switch_to.window(self.windows[0])
        self.closed_windows(0)
        time.sleep(1)
