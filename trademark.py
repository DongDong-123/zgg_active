import os
import time
import xlwt
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from readConfig import ReadConfig
from db import DbOperate
from Common import Common


class FunctionName(type):
    def __new__(cls, name, bases, attrs, *args, **kwargs):
        count = 0
        attrs["__Func__"] = []
        for k, v in attrs.items():
            if "trademark_" in k:
                attrs["__Func__"].append(k)
                count += 1

        attrs["__FuncCount__"] = count
        return type.__new__(cls, name, bases, attrs)

    def get_count(cls):
        pass


class Execute(object, metaclass=FunctionName):
    def __init__(self):
        self.common = Common()
        # 登录
        self.common.driver = self.common.driver
        # Excel写入
        self.row = 0
        self.workbook = xlwt.Workbook(encoding='utf-8')
        self.booksheet = self.workbook.add_sheet('Sheet1')
        self.timetemp = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())  # 存储Excel表格文件名编号
        # 每个案件的数量
        self.number = 1
        self.report_path = ReadConfig().save_report()
        self.dboperate = DbOperate()
        self.db = "case"
        self.catlog = 2

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

    # 立即申请
    def apply_now(self):
        self.common.driver.find_element_by_xpath("//div[@class='ui-zlsq-gwc']/a[1]").click()

    # 处理价格字符
    def process_price(self, price):
        if "￥" in price:
            price = price.replace("￥", '')
        return price

    # 国内商标
    def trademark_adviser_register(self):
        all_type = [u'专属顾问注册', u'专属加急注册', u'专属双享注册', u'专属担保注册']
        for trademark_type in all_type:
            if self.dboperate.is_member(self.db, trademark_type):
                try:
                    locator = (By.XPATH, "(.//div[@class='fl isnaMar'])[2]")
                    WebDriverWait(self.common.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
                    aa = self.common.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[2]")
                    ActionChains(self.common.driver).move_to_element(aa).perform()
                    self.common.driver.find_element_by_link_text(trademark_type).click()
                    # 切换至新窗口
                    self.common.windows = self.common.driver.window_handles
                    self.common.driver.switch_to_window(self.common.windows[-1])

                    self.apply_now()
                    # 切换至选择商标分类页面
                    self.common.windows = self.common.driver.window_handles
                    self.common.driver.switch_to_window(self.common.windows[-1])
                    while not self.common.driver.find_element_by_id("costesNum").is_displayed():
                        time.sleep(0.5)
                    # 获取详情页 价格
                    case_name = trademark_type
                    detail_price = self.common.driver.find_element_by_xpath("(.//div[@class='info-checkedtop']/p/span)").text
                    # detail_price = self.common.driver.find_element_by_xpath("(.//div[@class='bottomin']/p[1]/span)").text
                    # print("商标页价格", total_price)
                    detail_price = self.common.process_price(detail_price)

                    print("{}详情页价格".format(case_name), detail_price)
                    self.common.save_to_mysql([case_name, detail_price, self.catlog])

                    # 删除已执行的类型
                    self.dboperate.del_elem(self.db, trademark_type)
                    time.sleep(1)
                    self.common.closed_windows(0)
                except Exception as e:
                    print('错误信息', e)
                    self.common.driver.switch_to_window(self.common.windows[0])

    # 国际商标
    def trademark_international(self):
        all_type = [u'美国商标注册', u'日本商标注册', u'韩国商标注册', u'台湾商标注册', u'香港商标注册', u'德国商标注册',
                    u'欧盟商标注册', u'马德里国际商标', u'非洲知识产权组织']
        for international_type in all_type:
            if self.dboperate.is_member(self.db, international_type):
                # print(self.dboperate.is_member(international_type))
                try:
                    locator = (By.XPATH, ".//div[@class='isnav-first']/div[1]/h2")
                    WebDriverWait(self.common.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
                    aa = self.common.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[2]")
                    ActionChains(self.common.driver).move_to_element(aa).perform()
                    self.common.driver.find_element_by_link_text(international_type).click()
                    # 切换至新窗口
                    self.common.windows = self.common.driver.window_handles
                    self.common.driver.switch_to_window(self.common.windows[-1])
                    # 商标分类
                    self.common.driver.find_element_by_xpath("//a[@class='theme-fl']").click()
                    time.sleep(0.5)
                    self.common.driver.find_element_by_xpath("//ul[@class='theme-ul']/li[1]/p").click()
                    time.sleep(0.5)
                    self.common.driver.find_element_by_xpath("//div[@class='theme-btn']/a[3]").click()
                    time.sleep(0.5)
                    while not self.common.driver.find_element_by_id("totalfee").is_displayed():
                        time.sleep(0.5)
                    # 获取详情页 价格
                    case_name = str(international_type)
                    detail_price = self.common.driver.find_element_by_xpath(
                        "(.//div[@class='sames']//label[@id='totalfee'])").text
                    print("{}详情页价格".format(case_name), detail_price)
                    self.dboperate.del_elem(self.db, international_type)
                    time.sleep(1)
                    self.common.save_to_mysql([case_name, detail_price, self.catlog])
                    self.common.closed_windows(0)
                except Exception as e:
                    print(e)
                    self.common.driver.switch_to_window(self.common.windows[0])

    # 共用部分
    def trademark_famous_brand(self):
        all_type = [u'申请商标更正', u'出具商标注册证明申请', u'补发商标注册证申请', u'商标转让', u'商标注销', u'商标变更', u'商标诉讼', u'证明商标注册',
                    u'集体商标注册', u'驰名商标认定']
        for trademark in all_type:
            if self.dboperate.is_member(self.db, trademark):
                try:
                    locator = (By.XPATH, "(.//div[@class='fl isnaMar'])[2]")
                    WebDriverWait(self.common.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
                    aa = self.common.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[2]")
                    ActionChains(self.common.driver).move_to_element(aa).perform()
                    self.common.driver.find_element_by_link_text(trademark).click()
                    # 切换至新窗口
                    self.common.windows = self.common.driver.window_handles
                    self.common.driver.switch_to_window(self.common.windows[-1])
                    while not self.common.driver.find_element_by_id("totalfee").is_displayed():
                        time.sleep(0.5)
                    # 获取详情页 价格
                    case_name = str(trademark)
                    detail_price = self.common.driver.find_element_by_xpath("(.//div[@class='sames']//label[@id='totalfee'])").text
                    print("{}详情页价格".format(case_name), detail_price)
                    self.dboperate.del_elem(self.db, trademark)
                    self.common.save_to_mysql([case_name, detail_price, self.catlog])

                    time.sleep(1)
                    self.common.closed_windows(0)
                except Exception as e:
                    print('错误信息', e)
                    self.common.driver.switch_to_window(self.common.windows[0])

        time.sleep(1)

    # 商标驳回复审-（普通，双保）
    def trademark_ordinary_reject(self):
        this_type = u'商标驳回复审'
        if self.dboperate.is_member(self.db, this_type):
            locator = (By.XPATH, ".//div[@class='isnav-first']/div[1]/h2")
            WebDriverWait(self.common.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
            aa = self.common.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[2]")
            ActionChains(self.common.driver).move_to_element(aa).perform()
            self.common.driver.find_element_by_link_text(this_type).click()
            # 切换至新窗口
            self.common.windows = self.common.driver.window_handles
            self.common.driver.switch_to_window(self.common.windows[-1])
            # 服务类型
            for num in [2271, 22712]:
                try:
                    self.common.driver.find_element_by_xpath(".//ul[@id='ulType']/li[@pt='{}']/a".format(num)).click()
                    case_name = self.common.driver.find_element_by_xpath(
                        ".//ul[@id='ulType']/li[@pt='{}']/a".format(num)).text
                    case_name = "-".join([this_type, case_name])

                    # 数量加减
                    # self.common.number_add()
                    # # self.common.number_minus()
                    time.sleep(0.5)
                    while not self.common.driver.find_element_by_id("totalfee").is_displayed():
                        time.sleep(0.5)
                    # 获取详情页 价格
                    detail_price = self.common.driver.find_element_by_xpath("(.//div[@class='sames']//label[@id='totalfee'])").text
                    print("{}详情页价格".format(case_name), detail_price)
                    self.common.save_to_mysql([case_name, detail_price, self.catlog])
                except Exception as e:
                    print(e)
                    self.common.driver.switch_to_window(self.common.windows[1])
            self.common.closed_windows(0)
            self.dboperate.del_elem(self.db, this_type)
            time.sleep(1)

    # 商标异议 （异议申请、异议答辩）
    def trademark_objection_apply(self):
        this_type = u'商标异议'
        if self.dboperate.is_member(self.db, this_type):
            locator = (By.XPATH, ".//div[@class='isnav-first']/div[1]/h2")
            WebDriverWait(self.common.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
            aa = self.common.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[2]")
            ActionChains(self.common.driver).move_to_element(aa).perform()
            self.common.driver.find_element_by_link_text(this_type).click()
            # 切换至新窗口
            self.common.windows = self.common.driver.window_handles
            self.common.driver.switch_to_window(self.common.windows[-1])
            # 业务方向:异议申请、异议答辩、不予注册复审
            for num in [22721, 22722, 22723]:
                try:
                    self.common.driver.find_element_by_xpath(".//li[@pt='{}']/a".format(num)).click()
                    case_name = self.common.driver.find_element_by_xpath(".//li[@pt='{}']/a".format(num)).text
                    # 数量加减
                    # self.common.number_add()
                    # # self.common.number_minus()
                    case_name = "-".join([this_type, case_name])
                    while not self.common.driver.find_element_by_id("totalfee").is_displayed():
                        time.sleep(0.5)
                    # 获取详情页 价格
                    detail_price = self.common.driver.find_element_by_xpath("(.//div[@class='sames']//label[@id='totalfee'])").text
                    print("{}详情页价格".format(case_name), detail_price)
                    self.common.save_to_mysql([case_name, detail_price, self.catlog])

                except Exception as e:
                    print(e)
                    self.common.driver.switch_to_window(self.common.windows[1])
            self.common.closed_windows(0)
            self.dboperate.del_elem(self.db, this_type)
            time.sleep(1)

    # 商标撤三答辩--（商标撤三申请、商标撤三答辩）
    def trademark_brand_revoke_answer(self):
        this_type = u'商标撤销'
        if self.dboperate.is_member(self.db, this_type):
            locator = (By.XPATH, ".//div[@class='isnav-first']/div[1]/h2")
            WebDriverWait(self.common.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
            aa = self.common.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[2]")
            ActionChains(self.common.driver).move_to_element(aa).perform()
            self.common.driver.find_element_by_link_text(this_type).click()
            # 切换至新窗口
            self.common.windows = self.common.driver.window_handles
            self.common.driver.switch_to_window(self.common.windows[-1])
            # 业务方向:商标撤三申请、商标撤三答辩
            for num in range(1, 3):
                try:
                    self.common.driver.find_element_by_xpath(".//ul[@p='2273']/li[{}]/a".format(num)).click()
                    case_name = self.common.driver.find_element_by_xpath(".//ul[@p='2273']/li[{}]/a".format(num)).text
                    # 数量加减
                    # self.common.number_add()
                    # # self.common.number_minus()
                    case_name = "-".join([this_type, case_name])

                    time.sleep(0.5)
                    while not self.common.driver.find_element_by_id("totalfee").is_displayed():
                        time.sleep(0.5)
                    # 获取详情页 价格
                    detail_price = self.common.driver.find_element_by_xpath("(.//div[@class='sames']//label[@id='totalfee'])").text
                    print("{}详情页价格".format(case_name), detail_price)

                    self.common.save_to_mysql([case_name, detail_price, self.catlog])

                except Exception as e:
                    print(e)
                    self.common.driver.switch_to_window(self.common.windows[1])
            self.common.closed_windows(0)
            self.dboperate.del_elem(self.db, this_type)
            time.sleep(1)

    # 商标无效宣告--（商标无效宣告、商标无效宣告答辩）
    def trademark_brand_invalid_declare(self):
        this_type = u'商标无效'
        if self.dboperate.is_member(self.db, this_type):
            locator = (By.XPATH, ".//div[@class='isnav-first']/div[1]/h2")
            WebDriverWait(self.common.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
            aa = self.common.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[2]")
            ActionChains(self.common.driver).move_to_element(aa).perform()
            self.common.driver.find_element_by_link_text(this_type).click()
            # 切换至新窗口
            self.common.windows = self.common.driver.window_handles
            self.common.driver.switch_to_window(self.common.windows[-1])
            # 业务方向:商标无效宣告、商标无效宣告答辩
            for num in range(1, 3):
                try:
                    self.common.driver.find_element_by_xpath(".//ul[@p='2279']/li[{}]/a".format(num)).click()
                    case_name = self.common.driver.find_element_by_xpath(".//ul[@p='2279']/li[{}]/a".format(num)).text
                    case_name = "-".join([this_type, case_name])
                    # 数量加减
                    # self.common.number_add()
                    # # self.common.number_minus()

                    time.sleep(0.5)
                    while not self.common.driver.find_element_by_id("totalfee").is_displayed():
                        time.sleep(0.5)
                    # 获取详情页 价格
                    detail_price = self.common.driver.find_element_by_xpath("(.//div[@class='sames']//label[@id='totalfee'])").text
                    print("{}详情页价格".format(case_name), detail_price)
                    self.common.save_to_mysql([case_name, detail_price, self.catlog])

                except Exception as e:
                    print(e)
                    self.common.driver.switch_to_window(self.common.windows[1])
            self.common.closed_windows(0)
            self.dboperate.del_elem(self.db, this_type)
            time.sleep(1)

    # 商标续展--（续展申请、宽展申请、补发续展证明）
    def trademark_brand_extension_01(self):
        this_type = u'商标续展'
        if self.dboperate.is_member(self.db, this_type):
            locator = (By.XPATH, ".//div[@class='isnav-first']/div[1]/h2")
            WebDriverWait(self.common.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
            aa = self.common.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[2]")
            ActionChains(self.common.driver).move_to_element(aa).perform()
            self.common.driver.find_element_by_link_text(this_type).click()
            # 切换至新窗口
            self.common.windows = self.common.driver.window_handles
            self.common.driver.switch_to_window(self.common.windows[-1])
            # 业务方向:续展申请、宽展申请、补发续展证明
            for num in range(1, 4):
                try:
                    self.common.driver.find_element_by_xpath(".//ul[@p='2274']/li[{}]/a".format(num)).click()
                    case_name = self.common.driver.find_element_by_xpath(".//ul[@p='2274']/li[{}]/a".format(num)).text
                    case_name = "-".join([this_type, case_name])

                    # 数量加减
                    # self.common.number_add()
                    # # self.common.number_minus()
                    time.sleep(0.5)
                    while not self.common.driver.find_element_by_id("totalfee").is_displayed():
                        time.sleep(0.5)
                    # 获取详情页 价格
                    detail_price = self.common.driver.find_element_by_xpath("(.//div[@class='sames']//label[@id='totalfee'])").text
                    print("{}详情页价格".format(case_name), detail_price)
                    self.common.save_to_mysql([case_name, detail_price, self.catlog])

                except Exception as e:
                    print(e)
                    self.common.driver.switch_to_window(self.common.windows[1])
            self.common.closed_windows(0)
            self.dboperate.del_elem(self.db, this_type)
            time.sleep(1)

    # 商标许可备案 --(许可备案、变更（被）许可人名称、许可提前终止)
    def trademark_brand_permit(self):
        this_type = u'商标许可备案'
        if self.dboperate.is_member(self.db, this_type):
            locator = (By.XPATH, ".//div[@class='isnav-first']/div[1]/h2")
            WebDriverWait(self.common.driver, 30, 0.5).until(EC.element_to_be_clickable(locator))
            aa = self.common.driver.find_element_by_xpath("(.//div[@class='fl isnaMar'])[2]")
            ActionChains(self.common.driver).move_to_element(aa).perform()
            self.common.driver.find_element_by_link_text(this_type).click()
            # 切换至新窗口
            self.common.windows = self.common.driver.window_handles
            self.common.driver.switch_to_window(self.common.windows[-1])
            # 业务方向:许可备案、变更（被）许可人名称、许可提前终止
            for num in range(1, 4):
                try:
                    self.common.driver.find_element_by_xpath(".//ul[@p='2278']/li[{}]/a".format(num)).click()
                    case_name = self.common.driver.find_element_by_xpath(".//ul[@p='2278']/li[{}]/a".format(num)).text
                    case_name = "-".join([this_type, case_name])

                    # 数量加减
                    # self.common.number_add()
                    # # self.common.number_minus()
                    time.sleep(0.5)
                    while not self.common.driver.find_element_by_id("totalfee").is_displayed():
                        time.sleep(0.5)
                    # 获取详情页 价格
                    detail_price = self.common.driver.find_element_by_xpath("(.//div[@class='sames']//label[@id='totalfee'])").text
                    print("{}详情页价格".format(case_name), detail_price)
                    self.common.save_to_mysql([case_name, detail_price, self.catlog])

                except Exception as e:
                    print(e)
                    self.common.driver.switch_to_window(self.common.windows[1])
            self.common.closed_windows(0)
            self.dboperate.del_elem(self.db, this_type)
            time.sleep(1)
