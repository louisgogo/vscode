from selenium import webdriver
from time import sleep
from shutil import move
import os
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


class system():
    # 设定下载年月
    def __init__(self, n, a, year, month, skip, driver):
        # n表示报表序号，包括fastreport,budgetreport,addendum
        self.nolist = n
        self.start = a
        self.year = year
        self.month = month
        self.end = 38
        self.dir = {
            '1': '存款明细表',
            '2': '直接融资明细表',
            '4': '银行贷款明细表',
            '5': '其他贷款明细表'
        }
        self.dirt = {
            '1': '存款基本表',
            '2': '直接融资统计表',
            '3': '银行贷款统计表',
            '4': '其他贷款统计表'
        }
        self.skip = skip
        self.driver = driver

    # 登录软件界面
    def log_on(self):
        # 模拟登陆系统
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
            "//form[@id='login_form']/ul[@class='login-ul']/li[@class='clearfix']/a"
        )
        elem.click()
        # 进入报表编制界面
        elem = self.driver.find_element_by_xpath(
            "//div[@class='row rap-system-list']/div[1]/a/img")
        elem.click()
        elem = self.driver.find_elements_by_xpath("//li[@class='rap-module']")
        elem[1].click()
        sleep(1)
        elem = self.driver.find_element_by_xpath(
            "//li[@class='rap-module open']/ul/li[1]")
        elem.click()
        sleep(1)

    # 进入需要下载的报表界面
    def select(self):
        # 选择窗口
        self.driver.switch_to.frame(
            "rap-iframe-func-2c90e4df4c02ca32014c02f557d6002c")
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
        # 选择月份
        elem = self.driver.find_element_by_id("month")
        elem.click()
        sleep(2)
        elem = self.driver.find_element_by_xpath(
            "//select[@id='month']/option[@value=%s]" % self.month)
        elem.click()
        sleep(4)
        # 选择具体报表
        elem = self.driver.find_element_by_id("myTab")
        elem.click()
        sleep(2)
        elem = self.driver.find_element_by_xpath(
            "//ul[@id='menuList']/li[2]/a")
        ActionChains(self.driver).move_to_element(elem).perform()
        sleep(2)
        elem = self.driver.find_element_by_xpath(
            "//ul[@id='menuList']/li[2]/ul/li[3]/a")
        ActionChains(self.driver).move_to_element(elem).perform()
        sleep(2)
        ActionChains(self.driver).double_click(elem).perform()
        sleep(2)
        self.driver.switch_to.default_content()

    #下载合并报表
    def download_1(self):
        # 进入报表数据下载界面
        for i in range(self.start, self.end):
            if i not in self.skip:
                self.driver.switch_to.frame(
                    "rap-iframe-func-2c90e4df4c02ca32014c02f557d6002c")
                self.driver.switch_to.frame("selEpList")
                # 选择单位
                elem = self.driver.find_element_by_xpath(
                    "//ul[@id='treeLeft_1_ul']/li[%s]" % i)
                elem.click()
                # 控制滚动轴
                target = self.driver.find_element_by_xpath(
                    "//ul[@id='treeLeft_1_ul']/li[%s]" % i)
                self.driver.execute_script("arguments[0].scrollIntoView();",
                                           target)
                text = elem.text
                text = str(
                    i) + "号" + self.year + "年" + self.month + "月" + text + "-"
                sleep(1)
                self.driver.switch_to.parent_frame()
                self.driver.switch_to.frame("mainFrame")
                for j in self.nolist:
                    while True:
                        # 循环选取各个报表
                        try:
                            reportname = self.dir[str(j)]
                            sleep(1)
                            elem = self.driver.find_element_by_id("exportType")
                            elem.click()
                            elem = self.driver.find_element_by_xpath(
                                "//select[@id='exportType']/option[%s]" % j)
                            elem.click()
                            sleep(2)
                            print("准备下载")
                            if j == 5:
                                elem = self.driver.find_element_by_xpath(
                                    "//tr[@id='toolBar']/td[2]")
                            else:
                                elem = self.driver.find_element_by_xpath(
                                    "//tr[@id='toolBar']/td[3]")
                            elem.click()
                            sleep(1)
                            move("D:\\投资控股\\商业智能BI\\快报\\快报附表审核\\临时文件\\%s.xls" %
                                 reportname,
                                 "D:\\投资控股\\商业智能BI\\快报\\快报附表审核\\%s\\%s.xls" %
                                 (reportname, text + reportname))
                            break
                        except Exception as e:
                            print("文件下载失败：", text, e)
                            pass
                print("下载结束")
                sleep(1)
                self.driver.switch_to.default_content()
        self.driver.switch_to.frame(
            "rap-iframe-func-2c90e4df4c02ca32014c02f557d6002c")
        self.driver.switch_to.frame("selEpList")
        elem = self.driver.find_element_by_id("treeLeft_1_span")
        elem.click()
        self.driver.switch_to.default_content()

    def download_2(self):
        for j in range(1, 5):
            # 选择第一个非合并报表
            self.driver.switch_to.frame(
                "rap-iframe-func-2c90e4df4c02ca32014c02f557d6002c")
            self.driver.switch_to.frame("selEpList")
            elem = self.driver.find_element_by_xpath(
                "//ul[@id='treeLeft_1_ul']/li[%s]" % self.skip[0])
            elem.click()
            # 退回报表选择界面
            self.driver.switch_to.parent_frame()
            self.driver.switch_to.frame("mainFrame")
            # 选择具体报表
            elem = self.driver.find_element_by_id("myTab")
            elem.click()
            sleep(2)
            elem = self.driver.find_element_by_xpath(
                "//ul[@id='menuList']/li[2]/a")
            ActionChains(self.driver).move_to_element(elem).perform()
            sleep(2)
            elem = self.driver.find_element_by_xpath(
                "//ul[@id='menuList']/li[2]/ul/li[%s]/a" % (j + 2))
            ActionChains(self.driver).move_to_element(elem).perform()
            sleep(1)
            ActionChains(self.driver).double_click(elem).perform()
            sleep(2)
            for i in self.skip:
                self.driver.switch_to.parent_frame()
                self.driver.switch_to.frame("selEpList")
                # 选择单位
                elem = self.driver.find_element_by_xpath(
                    "//ul[@id='treeLeft_1_ul']/li[%s]" % i)
                elem.click()
                # 控制滚动轴
                target = self.driver.find_element_by_xpath(
                    "//ul[@id='treeLeft_1_ul']/li[%s]" % i)
                self.driver.execute_script("arguments[0].scrollIntoView();",
                                           target)
                text = elem.text
                text = str(
                    i) + "号" + self.year + "年" + self.month + "月" + text + "-"
                sleep(1)
                self.driver.switch_to.parent_frame()
                self.driver.switch_to.frame("mainFrame")
                while True:
                    # 循环选取各个报表
                    try:
                        if "统计表" in self.dirt[str(j)]:
                            reportname = self.dirt[str(j)].replace(
                                "统计表", "明细表")
                        else:
                            reportname = self.dirt[str(j)].replace(
                                "基本表", "明细表")
                        print(reportname)
                        sleep(1)
                        elem = self.driver.find_element_by_xpath(
                            "//td[@id='excelBtn']/button[1]")
                        elem.click()
                        sleep(1)
                        print("准备下载")
                        sleep(1)
                        move("D:\\投资控股\\商业智能BI\\快报\\快报附表审核\\临时文件\\%s.xls" %
                             self.dirt[str(j)],
                             "D:\\投资控股\\商业智能BI\\快报\\快报附表审核\\%s\\%s.xls" %
                             (reportname, text + self.dirt[str(j)]))
                        break
                    except Exception as e:
                        print("文件下载失败：", text, e)
                        pass
                print("下载结束")
                sleep(1)
            self.driver.switch_to.default_content()


if __name__ == '__main__':
    # 初始化浏览器
    options = webdriver.ChromeOptions()
    prefs = {
        'profile.default_co tent_settings.popups': 0,
        'download.default_directory': 'D:\\投资控股\\商业智能BI\\快报\\快报附表审核\\临时文件'
    }
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(chrome_options=options)
    # n表示附表编号列表，{1:'存款明细表',2:'直接融资明细表',4:'银行贷款明细表',5:'其他贷款明细表'}
    # a表示初始编号,skip参数为列表
    a = 1
    n = [1, 2, 4, 5]
    year = input("请输入年份")
    month = input("请输入月份")
    skip = [5, 6, 19, 22, 23, 26, 27, 30, 36]

    test = system(n, a, year, month, skip, driver)
    test.log_on()
    test.select()
    test.download_1()
    test.select()
    test.download_2()