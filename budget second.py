from selenium import webdriver
from time import sleep
from shutil import move
import os


class budget():
    # 设定下载年月
    def __init__(self, t, a, year, driver, switch):
        self.type = t
        self.num = a
        self.year = year
        self.driver = driver
        self.adjust = switch
        self.end = 38

    def log_on(self):
        self.driver.implicitly_wait(20)
        self.driver.get('https://sam.szgzw.gov.cn:9443/szgzw/login')
        self.driver.maximize_window()
        sleep(2)
        elem = self.driver.find_element_by_xpath(
            "//div[@id='userNumBox']/input[@id='usernumber']")
        elem.send_keys("tkkh_ys5")
        elem = self.driver.find_element_by_xpath(
            "//div[@id='pwdBox']/input[@id='mm']")
        elem.send_keys("4gR5Nbzp86Vu")
        elem = self.driver.find_element_by_xpath(
            "//div[@id='valitxt']/input[@id='sendsms']")
        elem.click()
        sendsms = input("请输入验证码")
        elem = self.driver.find_element_by_xpath(
            "//div[@id='valiBox']/input[@id='valiCode']")
        elem.send_keys(sendsms)
        sleep(2)
        elem = self.driver.find_element_by_xpath(
            "//form[@id='login_form']/ul[@class='login-ul']/li[@class='clearfix']/a")
        elem.click()
        # 进入报表编制界面
        elem = self.driver.find_element_by_xpath(
            "//div[@class='row rap-system-list']/div[1]/a/img")
        elem.click()
        elem = self.driver.find_elements_by_xpath("//li[@class='rap-module']")
        elem[3].click()
        sleep(1)
        elem = self.driver.find_element_by_xpath(
            "//li[@class='rap-module open']/ul/li[2]")
        elem.click()
        sleep(1)

    def select(self):
        # 选择窗口
        sleep(3)
        self.driver.switch_to.frame(
            "rap-iframe-func-2c90e4d74a38343d014a38cf1dcf009a")
        self.driver.switch_to.frame("selEpList")
        self.driver.switch_to.parent_frame()
        self.driver.switch_to.frame("mainFrame")
        sleep(4)
        # 选择年份
        elem = self.driver.find_element_by_id("year")
        elem.click()
        sleep(2)
        elem = self.driver.find_element_by_xpath(
            "//select[@id='year']/option[@value=%s]" % self.year)
        elem.click()
        sleep(1)
        # 选择需要下载的报表类型
        elem = self.driver.find_elements_by_xpath(
            "//td[@id='menuItem2']")
        self.driver.execute_script(
            "arguments[0].scrollIntoView();", elem[self.type])
        elem[self.type].click()
        ReportType1 = elem[self.type].text[2:]
        ReportType1 = ReportType1.replace("、", "")
        ReportType2 = ReportType1.replace("汇总", "")
        self.driver.switch_to.default_content()
        sleep(1)
        return ReportType1, ReportType2

    def down_loads(self, ReportType1, ReportType2):
        # 进入报表数据下载界面
        for i in range(self.num, self.end):
            sleep(1)
            self.driver.switch_to.frame(
                "rap-iframe-func-2c90e4d74a38343d014a38cf1dcf009a")
            self.driver.switch_to.frame("selEpList")
            # 选择单位
            elem = self.driver.find_element_by_xpath("//ul[@id='treeLeft_1_ul']/li[%s]" % i)
            elem.click()
            # 控制滚动轴
            target = self.driver.find_element_by_xpath("//ul[@id='treeLeft_1_ul']/li[%s]" % i)
            self.driver.execute_script(
                "arguments[0].scrollIntoView();", target)
            text = elem.text
            text = str(i) + "号" + year + "年" + text + "-" + ReportType2
            sleep(1)
            self.driver.switch_to.parent_frame()
            self.driver.switch_to.frame("mainFrame")
            while True:
                try:
                    sleep(1)
                    print("准备下载")
                    elem = self.driver.find_element_by_id("excelBtn")
                    elem.click()
                    break
                except:
                #    driver.refresh()
                    pass
            print("下载结束")
            sleep(1)
            while True:
                try:
                    print("数据1", ReportType1)
                    sleep(1)
                    move("D:\投资控股\商业智能BI\系统数据\当月数据\年度预算\%s.xls" % ReportType1,
                         "D:\投资控股\商业智能BI\系统数据\当月数据\年度预算\%s.xls" % text)
                    print("文件生成完毕")
                    break
                except:
                    print("数据2", ReportType2)
                    try:
                        move("D:\投资控股\商业智能BI\系统数据\当月数据\年度预算\%s.xls" % ReportType2,
                             "D:\投资控股\商业智能BI\系统数据\当月数据\年度预算\%s.xls" % text)
                        print("文件生成完毕")
                        break
                    except:
                        pass
            self.driver.switch_to.default_content()
        self.driver.switch_to.frame(
            "rap-iframe-func-2c90e4d74a38343d014a38cf1dcf009a")
        self.driver.switch_to.frame("selEpList")
        elem = self.driver.find_element_by_id("treeLeft_1_span")
        elem.click()
        self.driver.switch_to.default_content()

    def exit_out(self):
        os._exit(0)


# t = int(input("抓取报告类型（eg:3-资产负债/4-利润表/6-现金流量/7-现金流量续/16-期间费用/17-期间费用续）"))
t = [4,16,17]
a = int(input("开始抓取的公司编号（eg:1）"))
year = input("请输入年份(eg:2018)")
count = 0
# 设置下载方式
options = webdriver.ChromeOptions()
prefs = {'profile.default_content_settings.popups': 0,
         'download.default_directory': 'D:\投资控股\商业智能BI\系统数据\当月数据\年度预算'}
options.add_experimental_option('prefs', prefs)
switch = False
# 模拟登陆界面
driver = webdriver.Chrome(chrome_options=options)
for i in t:
    bud = budget(i, a, year, driver, switch)
    if count == 0:
        bud.log_on()
        count += 1
    ReportType1, ReportType2 = bud.select()
    bud.down_loads(ReportType1, ReportType2)
    switch = True
bud.exit_out()
