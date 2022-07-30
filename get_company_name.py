from selenium import webdriver
import xlrd
import xlwt
from time import sleep, time

# ========Read xls========

book1 = xlrd.open_workbook('财务指标_代码.xlsx')
sheet1 = book1.sheets()[0]

# ========Get data========

opt = webdriver.ChromeOptions()
opt.headless = False
driver = webdriver.Chrome(options=opt)
driver.maximize_window()

wait_time = 5

saved_num = 0
with open('company_name.txt', 'a+') as f:
    pass

with open('company_name.txt', 'r') as f:
    saved_num = len(f.readlines())

driver.get('http://so.eastmoney.com/web/s?keyword=000001&pageindex=1')
# wait = input('start ?')

for i in range(saved_num, sheet1.nrows):

    saved_msg = ''

    company_name = str(sheet1.cell(i, 0).value)

    print(company_name, "  start : ")

    with open('company_name.txt', 'a+') as f:
        saved_msg += company_name

        driver.get('http://so.eastmoney.com/web/s?keyword=' + company_name + '&pageindex=1')
        time_start = time()
        success = False
        while not success:
            try:
                driver.find_element_by_xpath('//div[@class="amodule"]/h3/a').click()
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

# # ========Write xls========

workbook = xlwt.Workbook()

worksheet = workbook.add_sheet('company_name')

for i in range(sheet1.nrows):
    for j in range(sheet1.ncols):
        worksheet.write(i, j, sheet1.cell(i, j).value)

with open('company_name.txt', 'r') as f:
    lines = f.readlines()

    for i in range(sheet1.nrows):
        for line in lines:
            if line[:6] == sheet1.cell(i, 0).value:
                worksheet.write(i, 1, line[6:])
                break

workbook.save('company_name.xls')