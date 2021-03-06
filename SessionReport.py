from selenium import webdriver
from time import sleep
from os import rename
from shutil import move
from selenium.webdriver.common.keys import Keys
import os

# 设定下载年月
t = int(input("抓取报告类型（eg:0-资产负债表/1-利润表/2-现金流量）"))
a = int(input("开始抓取的公司编号（eg:1）"))
year = input("请输入年份(eg:2017)")
month = input("请输入月份（eg:03/06/09/12）")
# 设置下载方式
options = webdriver.ChromeOptions()
prefs = {'profile.default_content_settings.popups': 0,
         'download.default_directory': 'D:\投资控股\商业智能BI\Python\数据'}
options.add_experimental_option('prefs', prefs)
# 模拟登陆界面
driver = webdriver.Chrome(chrome_options=options)
driver.implicitly_wait(20)
driver.get('https://sam.szgzw.gov.cn:9443/szgzw/login')
driver.maximize_window()
sleep(2)
elem = driver.find_element_by_xpath(
    "//div[@class='loginIpt loginIpt-focus']/input[@id='usernumber']")
elem.send_keys("tkkh_ys5")
elem = driver.find_element_by_xpath("//div[@id='pwdBox']/input[@id='mm']")
elem.send_keys("4gR5Nbzp86Vu")
elem = driver.find_element_by_xpath(
    "//div[@id='valitxt']/input[@id='sendsms']")
elem.click()
sendsms = input("请输入验证码")
elem = driver.find_element_by_xpath(
    "//div[@id='valiBox']/input[@id='valiCode']")
elem.send_keys(sendsms)
sleep(2)
elem = driver.find_element_by_xpath(
    "//form[@id='login_form']/ul[@class='login-ul']/li[@class='clearfix']/a")
elem.click()
# 进入报表编制界面
elem = driver.find_element_by_xpath(
    "//div[@class='row rap-system-list']/div[1]/a/img")
elem.click()
elem = driver.find_elements_by_xpath("//li[@class='rap-module']")
elem[2].click()
sleep(1)
elem = driver.find_element_by_xpath(
    "//li[@class='rap-module open']/ul/li[1]")
elem.click()
sleep(4)
# 选择窗口
driver.switch_to.frame("rap-iframe-func-2c90e4df4c02ca32014c02f9285c0036")
driver.switch_to.frame("selEpList")
driver.switch_to.parent_frame()
driver.switch_to.frame("mainFrame")
sleep(4)
# 选择年份
elem = driver.find_element_by_id("year")
elem.click()
sleep(2)
elem = driver.find_element_by_xpath(
    "//select[@id='year']/option[@value=%s]" % year)
elem.click()
sleep(2)
# 选择月份
elem = driver.find_element_by_id("month")
elem.click()
sleep(2)
elem = driver.find_element_by_xpath(
    "//select[@id='month']/option[@value=%s]" % month)
elem.click()
sleep(1)
# 选择需要下载的报表类型
elem = driver.find_elements_by_xpath(
    "//td[@id='menuItem2']")
driver.execute_script("arguments[0].scrollIntoView();", elem[t])
elem[t].click()
ReportType1 = elem[t].text[2:]
ReportType1 = ReportType1.replace("、", "")
ReportType2 = ReportType1.replace("汇总", "")
driver.switch_to.default_content()
sleep(1)
# 进入报表数据下载界面
for i in range(a, 383):
    sleep(1)
    driver.switch_to.frame("rap-iframe-func-2c90e4df4c02ca32014c02f9285c0036")
    driver.switch_to.frame("selEpList")
    # 选择单位
    if i >= 2:
        elem = driver.find_element_by_id("treeLeft_%s_switch" % i)
        elem.click()
    elem = driver.find_element_by_id("treeLeft_%s_span" % i)
    elem.click()
    # 控制滚动轴
    target = driver.find_element_by_id("treeLeft_%s_span" % i)
    driver.execute_script("arguments[0].scrollIntoView();", target)
    text = elem.text
    text = str(i) + "号" + year + "年" + month + "月" + text + "-" + ReportType2
    sleep(2)
    print("准备下载")
    check = "初始化数据"
    driver.switch_to.parent_frame()
    driver.switch_to.frame("mainFrame")
    sleep(1)
    while check == "初始化数据":
        checkelem = driver.find_element_by_id("statusId")
        check = checkelem.text
        sleep(1)
    print(check)
    if check == "已审核" or check == "审核中":
        while True:
            try:
                sleep(1)
                elem = driver.find_element_by_xpath(
                    "//td[@id='excelBtn']/button")
                elem.click()
                break
            except:
                pass
        print("下载结束")
        sleep(2)
        while True:
            try:
                print("数据1", ReportType1)
                sleep(1)
                move("D:\投资控股\商业智能BI\Python\数据\%s.xls" % ReportType1,
                     "D:\投资控股\商业智能BI\Python\数据\%s.xls" % text)
                print("文件生成完毕")
                break
            except:
                print("数据2", ReportType2)
                try:
                    move("D:\投资控股\商业智能BI\Python\数据\%s.xls" % ReportType2,
                         "D:\投资控股\商业智能BI\Python\数据\%s.xls" % text)
                    print("文件生成完毕")
                    break
                except:
                    pass
    driver.switch_to.parent_frame()
    driver.switch_to.parent_frame()

os._exit(0)
