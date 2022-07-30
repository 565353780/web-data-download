from selenium import webdriver
import xlrd
import xlwt
from time import sleep, time

# from win32com import client
# import time
# import random
# from lxml import etree
# dirver = client.DispatchEx("InternetExplorer.Application")
# dirver.Navigate('http://sbj.saic.gov.cn/sbcx/')
# dirver.Visible = 1
# time.sleep(random.randint(2, 8))
# dirver.Document.body.getElementsByTagName("p")[3].firstElementChild.click()
# dirver.Visible = 1
# time.sleep(random.randint(8, 12))
# dirver.Document.body.getElementsByTagName("tbody")[1].click()
# time.sleep(random.randint(10, 20))
# for i in dirver.Document.body.getElementsByTagName("input"):
#     if i.name == 'request:hnc':
#         i.value = '百度'
# # 点击查询
# time.sleep(3)
# dirver.Visible = 1
# for i in dirver.Document.body.getElementsByTagName("input"):
#     if i.id == '_searchButton':
#         i.click()

# time.sleep(20)
# form_str=dirver.Document.body.getElementsByTagName("form")[2].innerHTML
# print(form_str)
# html_str = etree.HTML(form_str)
# tr_list = html_str.xpath('//tr[@class="ng-repeat"]')
# for tr in tr_list:
#     item = {}
#     item['注册号'] = tr.xpath('.//td[2]/a/text()')
#     item['国际分类'] = tr.xpath('.//td[3]/text()')
#     item['申请日期'] = tr.xpath('.//td[4]/text()')
#     item['商标名称'] = tr.xpath('.//td[5]/a/text()')
#     item['申请人名称'] = tr.xpath('.//td[6]/a/text()')

#     print(item)
#     with open('item.txt', 'w', encoding='utf-8') as f:
#         f.write(str(item))

# ========Read xls========

book1 = xlrd.open_workbook('company_name.xls')
sheet1 = book1.sheets()[0]

# ========Get data========

opt = webdriver.ChromeOptions()
opt.headless = False
driver = webdriver.Chrome(options=opt)
driver.maximize_window()

wait_time = 5

saved_num = 0
with open('company_msg.txt', 'a+') as f:
    pass

with open('company_msg.txt', 'r') as f:
    saved_num = len(f.readlines())

driver.get('http://wcjs.sbj.cnipa.gov.cn/txnS02.do?locale=zh_CN&kmcmNx0Q=qAcKqqkaZHkaZHkaZkBePYxjANmf1vQeKYeN19s.acgqqHq')
# wait = input('start ?')

for i in range(saved_num, sheet1.nrows):

    saved_msg = ''

    company_code = str(sheet1.cell(i, 0).value)

    company_name = str(sheet1.cell(i, 1).value)

    if company_name == '':
        continue

    print(company_name, "  start : ")

    with open('company_msg.txt', 'a+') as f:
        saved_msg += company_code + '|\|'

        driver.get('http://wcjs.sbj.cnipa.gov.cn/txnS02.do?locale=zh_CN&kmcmNx0Q=qAcKqqkaZHkaZHkaZkBePYxjANmf1vQeKYeN19s.acgqqHq')
        time_start = time()
        success = False
        while not success:
            try:
                text_edit = driver.find_element_by_xpath('//*[@id="submitForm"]/div/div[1]/table/tbody/tr[4]/td[2]/div/input')
                text_edit.clear()
                text_edit.send_keys(company_name)
                driver.find_element_by_xpath('//*[@id="_searchButton"]').click()
                success = True
            except:
                success = False
            
            if time() - time_start > wait_time:
                break
        
        if not success:
            saved_msg += '\n'
            f.write(saved_msg)
            continue
        
        driver.switch_to.window(driver.window_handles[1])

        time_start = time()
        success = False
        b = None
        while not success:
            try:
                b = driver.find_element_by_xpath('//div[@class="profile"]')
                success = True
            except:
                success = False
            
            if time() - time_start > wait_time:
                break
        
        if not success:
            saved_msg += '\n'
            f.write(saved_msg)
            continue
        
        print(b.text)
        saved_msg += str(b.text) + '\n'
        f.write(saved_msg)

        driver.close()

        driver.switch_to.window(driver.window_handles[0])

        print(str(i) + ' / ' + str(sheet1.nrows - 1))