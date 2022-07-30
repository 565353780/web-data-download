from selenium import webdriver
import xlrd
import xlwt

# ========Read xls========

book1 = xlrd.open_workbook('test1.xls')
sheet1 = book1.sheets()[0]



book2 = xlrd.open_workbook('test2.xls')
sheet2 = book2.sheets()[0]

# ========Get data========

opt = webdriver.ChromeOptions()
opt.headless = False
driver = webdriver.Chrome(options=opt)
driver.maximize_window()

# # ========================

# total_wait_episode = 100

# saved_num = 0
# with open('data_get1.txt', 'a+') as f:
#     pass

# with open('data_get1.txt', 'r') as f:
#     saved_num = len(f.readlines())

# for i in range(saved_num + 1, sheet1.nrows):
#     with open('data_get1.txt', 'a+') as f:
#         company_name = str(int(sheet1.cell(i, 0).value))

#         if company_name[0] == '6' and len(company_name) == 6:
#             print('上证模式')
#             print(company_name)

#             driver.get('http://www.sse.com.cn/home/search/?webswd=' + company_name + '%202018年年度报告')

#             name = None
#             name_got = False
#             get_name_episode = 0

#             while not name_got and get_name_episode < total_wait_episode:
#                 if get_name_episode > 0:
#                     print('\rRetrying getting company name : ' + str(get_name_episode), end='')
#                 try:
#                     name = driver.find_element_by_xpath('//p[@class="sse_query_title_left"]//i[@class="ssedesc"]').text
#                     name_got = True
#                 except:
#                     get_name_episode += 1
#                     pass

#             if get_name_episode == total_wait_episode:
#                 continue

#             while len(name) == 0 and get_name_episode < total_wait_episode:
#                 if get_name_episode > 0:
#                     print('\rRetrying getting company name : ' + str(get_name_episode), end='')
#                 get_name_episode += 1
#                 name = driver.find_element_by_xpath('//p[@class="sse_query_title_left"]//i[@class="ssedesc"]').text

#             if get_name_episode > 0:
#                 print('')

#             if get_name_episode == total_wait_episode:
#                 continue

#             print(name)

#             pdf = None
#             pdf_loaded = False
#             read_pdf_episode = 0

#             while not pdf_loaded and read_pdf_episode < total_wait_episode:
#                 if read_pdf_episode > 0:
#                     print('\rRetrying loading pdf : ' + str(read_pdf_episode), end='')
#                 try:
#                     pdf = driver.find_element_by_xpath('//a[@title="' + name + '2018年年度报告"]')
#                     pdf_loaded = True
#                 except:
#                     read_pdf_episode += 1
#                     pass

#             if read_pdf_episode > 0:
#                 print('')

#             if read_pdf_episode == total_wait_episode:
#                 continue

#             try:
#                 print(pdf.get_attribute('href'))

#                 f.write(pdf.get_attribute('href') + '\n')
#             except:
#                 continue
#         else:
#             print('深证模式')

#             name_len = len(company_name)

#             total_company_name = ''

#             for ii in range(6 - name_len):
#                 total_company_name += '0'

#             total_company_name += company_name

#             print(total_company_name)

#             driver.get('http://www.szse.cn/application/search/index.html?keyword=' + total_company_name)

#             name = None
#             name_got = False
#             get_name_episode = 0

#             while not name_got and get_name_episode < total_wait_episode:
#                 if get_name_episode > 0:
#                     print('\rRetrying getting company name : ' + str(get_name_episode), end='')
#                 try:
#                     name = driver.find_element_by_xpath('//span[@id="secChenkName"]').text
#                     name_got = True
#                 except:
#                     get_name_episode += 1
#                     pass

#             if get_name_episode == total_wait_episode:
#                 continue

#             while len(name) == 0 and get_name_episode < total_wait_episode:
#                 if get_name_episode > 0:
#                     print('\rRetrying getting company name : ' + str(get_name_episode), end='')
#                 name = driver.find_element_by_xpath('//span[@id="secChenkName"]').text
#                 get_name_episode += 1

#             if get_name_episode > 0:
#                 print('')

#             if get_name_episode == total_wait_episode:
#                 continue

#             if '*' in name:
#                 name = name.split('*')[1]

#             print(name)

#             driver.get('http://www.szse.cn/application/search/index.html?keyword=' + total_company_name + '%202018年年度报告')

#             opt = None
#             opt_changed = False
#             change_opt_episode = 0

#             while not opt_changed and change_opt_episode < total_wait_episode:
#                 if change_opt_episode > 0:
#                     print('\rRetrying changing option : ' + str(change_opt_episode), end='')
#                 try:
#                     opt = driver.find_element_by_xpath('//span[@class="glyphicon glyphicon-menu-down"]')
#                     opt_changed = True
#                 except:
#                     change_opt_episode += 1
#                     pass

#             if change_opt_episode > 0:
#                 print('')

#             if change_opt_episode == total_wait_episode:
#                 continue

#             opt.click()

#             opt = None
#             opt_changed = False
#             change_opt_episode = 0

#             while not opt_changed and change_opt_episode < total_wait_episode:
#                 if change_opt_episode > 0:
#                     print('\rRetrying changing option : ' + str(change_opt_episode), end='')
#                 try:
#                     opt = driver.find_element_by_xpath('//a[@title="标题+正文"]')
#                     opt_changed = True
#                 except:
#                     change_opt_episode += 1
#                     pass

#             if change_opt_episode > 0:
#                 print('')

#             if change_opt_episode == total_wait_episode:
#                 continue

#             opt.click()

#             search = None
#             search_changed = False
#             change_search_episode = 0

#             while not search_changed and change_search_episode < total_wait_episode:
#                 if change_search_episode > 0:
#                     print('\rRetrying changing search : ' + str(change_search_episode), end='')
#                 try:
#                     search = driver.find_element_by_xpath(
#                         '//div[@class="search-wrap sh-searchhint-container"]//button[@class="search-btn"]')
#                     search_changed = True
#                 except:
#                     change_search_episode += 1
#                     pass

#             if change_search_episode > 0:
#                 print('')

#             if change_search_episode == total_wait_episode:
#                 continue

#             search.click()

#             label_name = name + '：2018年年度报告'

#             label_list = None
#             label_list_changed = False
#             change_label_list_episode = 0

#             while not label_list_changed and change_label_list_episode < total_wait_episode:
#                 if change_label_list_episode > 0:
#                     print('\rRetrying getting label list : ' + str(change_label_list_episode), end='')
#                 try:
#                     label_list = driver.find_elements_by_xpath('//a[@class="text ellipsis pdf"]')
#                     label_list_changed = True
#                 except:
#                     change_label_list_episode += 1
#                     pass

#             if change_label_list_episode > 0:
#                 print('')

#             if change_label_list_episode == total_wait_episode:
#                 continue

#             file_found = False
#             file_label = None
#             get_file_label_episode = 0

#             while not file_found and get_file_label_episode < total_wait_episode:
#                 if get_file_label_episode > 0:
#                     print('\rRetrying getting file label : ' + str(get_file_label_episode), end='')
#                 try:
#                     label_list = driver.find_elements_by_xpath('//a[@class="text ellipsis pdf"]')
#                     for label in label_list:
#                         if (label_name in label.text) and ('摘要' not in label.text) and ('取消' not in label.text):
#                             file_label = label
#                             file_found = True
#                             break
#                 except:
#                     pass
#                 get_file_label_episode += 1

#             if get_file_label_episode > 0:
#                 print('')

#             if get_file_label_episode == total_wait_episode:
#                 continue

#             try:
#                 print(file_label.get_attribute('data-urlpath'))

#                 f.write(file_label.get_attribute('data-urlpath') + '\n')
#             except:
#                 continue

#     print(str(i) + ' / ' + str(sheet1.nrows - 1))



# saved_num = 0
# with open('data_get2.txt', 'a+') as f:
#     pass

# with open('data_get2.txt', 'r') as f:
#     saved_num = len(f.readlines())

# for i in range(saved_num + 1, sheet2.nrows):
#     with open('data_get2.txt', 'a+') as f:
#         company_name = str(int(sheet2.cell(i, 0).value))

#         if company_name[0] == '6' and len(company_name) == 6:
#             print('上证模式')
#             print(company_name)

#             driver.get('http://www.sse.com.cn/home/search/?webswd=' + company_name + '%202018年年度报告')

#             name = None
#             name_got = False
#             get_name_episode = 0

#             while not name_got and get_name_episode < total_wait_episode:
#                 if get_name_episode > 0:
#                     print('\rRetrying getting company name : ' + str(get_name_episode), end='')
#                 try:
#                     name = driver.find_element_by_xpath('//p[@class="sse_query_title_left"]//i[@class="ssedesc"]').text
#                     name_got = True
#                 except:
#                     get_name_episode += 1
#                     pass

#             if get_name_episode == total_wait_episode:
#                 continue

#             while len(name) == 0 and get_name_episode < total_wait_episode:
#                 if get_name_episode > 0:
#                     print('\rRetrying getting company name : ' + str(get_name_episode), end='')
#                 get_name_episode += 1
#                 name = driver.find_element_by_xpath('//p[@class="sse_query_title_left"]//i[@class="ssedesc"]').text

#             if get_name_episode > 0:
#                 print('')

#             if get_name_episode == total_wait_episode:
#                 continue

#             print(name)

#             pdf = None
#             pdf_loaded = False
#             read_pdf_episode = 0

#             while not pdf_loaded and read_pdf_episode < total_wait_episode:
#                 if read_pdf_episode > 0:
#                     print('\rRetrying loading pdf : ' + str(read_pdf_episode), end='')
#                 try:
#                     pdf = driver.find_element_by_xpath('//a[@title="' + name + '2018年年度报告"]')
#                     pdf_loaded = True
#                 except:
#                     read_pdf_episode += 1
#                     pass

#             if read_pdf_episode > 0:
#                 print('')

#             if read_pdf_episode == total_wait_episode:
#                 continue

#             print(pdf.get_attribute('href'))

#             f.write(pdf.get_attribute('href') + '\n')
#         else:
#             print('深证模式')

#             name_len = len(company_name)

#             total_company_name = ''

#             for ii in range(6 - name_len):
#                 total_company_name += '0'

#             total_company_name += company_name

#             print(total_company_name)

#             driver.get('http://www.szse.cn/application/search/index.html?keyword=' + total_company_name)

#             name = None
#             name_got = False
#             get_name_episode = 0

#             while not name_got and get_name_episode < total_wait_episode:
#                 if get_name_episode > 0:
#                     print('\rRetrying getting company name : ' + str(get_name_episode), end='')
#                 try:
#                     name = driver.find_element_by_xpath('//span[@id="secChenkName"]').text
#                     name_got = True
#                 except:
#                     get_name_episode += 1
#                     pass

#             if get_name_episode == total_wait_episode:
#                 continue

#             while len(name) == 0 and get_name_episode < total_wait_episode:
#                 if get_name_episode > 0:
#                     print('\rRetrying getting company name : ' + str(get_name_episode), end='')
#                 name = driver.find_element_by_xpath('//span[@id="secChenkName"]').text
#                 get_name_episode += 1

#             if get_name_episode > 0:
#                 print('')

#             if get_name_episode == total_wait_episode:
#                 continue

#             if '*' in name:
#                 name = name.split('*')[1]

#             print(name)

#             driver.get('http://www.szse.cn/application/search/index.html?keyword=' + total_company_name + '%202018年年度报告')

#             opt = None
#             opt_changed = False
#             change_opt_episode = 0

#             while not opt_changed and change_opt_episode < total_wait_episode:
#                 if change_opt_episode > 0:
#                     print('\rRetrying changing option : ' + str(change_opt_episode), end='')
#                 try:
#                     opt = driver.find_element_by_xpath('//span[@class="glyphicon glyphicon-menu-down"]')
#                     opt_changed = True
#                 except:
#                     change_opt_episode += 1
#                     pass

#             if change_opt_episode > 0:
#                 print('')

#             if change_opt_episode == total_wait_episode:
#                 continue

#             opt.click()

#             opt = None
#             opt_changed = False
#             change_opt_episode = 0

#             while not opt_changed and change_opt_episode < total_wait_episode:
#                 if change_opt_episode > 0:
#                     print('\rRetrying changing option : ' + str(change_opt_episode), end='')
#                 try:
#                     opt = driver.find_element_by_xpath('//a[@title="标题+正文"]')
#                     opt_changed = True
#                 except:
#                     change_opt_episode += 1
#                     pass

#             if change_opt_episode > 0:
#                 print('')

#             if change_opt_episode == total_wait_episode:
#                 continue

#             opt.click()

#             search = None
#             search_changed = False
#             change_search_episode = 0

#             while not search_changed and change_search_episode < total_wait_episode:
#                 if change_search_episode > 0:
#                     print('\rRetrying changing search : ' + str(change_search_episode), end='')
#                 try:
#                     search = driver.find_element_by_xpath(
#                         '//div[@class="search-wrap sh-searchhint-container"]//button[@class="search-btn"]')
#                     search_changed = True
#                 except:
#                     change_search_episode += 1
#                     pass

#             if change_search_episode > 0:
#                 print('')

#             if change_search_episode == total_wait_episode:
#                 continue

#             search.click()

#             label_name = name + '：2018年年度报告'

#             label_list = None
#             label_list_changed = False
#             change_label_list_episode = 0

#             while not label_list_changed and change_label_list_episode < total_wait_episode:
#                 if change_label_list_episode > 0:
#                     print('\rRetrying getting label list : ' + str(change_label_list_episode), end='')
#                 try:
#                     label_list = driver.find_elements_by_xpath('//a[@class="text ellipsis pdf"]')
#                     label_list_changed = True
#                 except:
#                     change_label_list_episode += 1
#                     pass

#             if change_label_list_episode > 0:
#                 print('')

#             if change_label_list_episode == total_wait_episode:
#                 continue

#             file_found = False
#             file_label = None
#             get_file_label_episode = 0

#             while not file_found and get_file_label_episode < total_wait_episode:
#                 if get_file_label_episode > 0:
#                     print('\rRetrying getting file label : ' + str(get_file_label_episode), end='')
#                 try:
#                     label_list = driver.find_elements_by_xpath('//a[@class="text ellipsis pdf"]')
#                     for label in label_list:
#                         if (label_name in label.text) and ('摘要' not in label.text) and ('取消' not in label.text):
#                             file_label = label
#                             file_found = True
#                             break
#                 except:
#                     pass
#                 get_file_label_episode += 1

#             if get_file_label_episode > 0:
#                 print('')

#             if get_file_label_episode == total_wait_episode:
#                 continue

#             print(file_label.get_attribute('data-urlpath'))

#             f.write(file_label.get_attribute('data-urlpath') + '\n')

#     print(str(i) + ' / ' + str(sheet2.nrows - 1))

# ========================

saved_num = 0
with open('data1_get.txt', 'a+') as f:
    pass

with open('data1_get.txt', 'r') as f:
    saved_num = len(f.readlines())

driver.get('https://www.tianyancha.com/')
wait = input('start ?')

for i in range(saved_num + 1, sheet1.nrows):

    company_name = str(int(sheet1.cell(i, 0).value))
    if company_name[0] == '6' and len(company_name) == 6:
        continue

    with open('data1_get.txt', 'a+') as f:
        company_name = sheet1.cell(i, 1).value
        driver.get('https://www.tianyancha.com/search?key=' + company_name)
        try:
            driver.find_elements_by_xpath('//a[@tyc-event-ch="CompanySearch.Company"]')[0].click()
            print('open new tab')
        except:
            print('no company for select')
            wait = input('start ?')
            try:
                driver.find_elements_by_xpath('//a[@tyc-event-ch="CompanySearch.Company"]')[0].click()
            except:
                print('no company for select checked')
                f.write('暂无信息||||暂无信息\n')
                continue

        driver.close()

        driver.switch_to.window(driver.window_handles[0])

        try:
            b = driver.find_element_by_xpath('//*[@id="company_web_top"]/div[2]/div[3]/div[1]/h1')
        except:
            print('company name not found')
            wait = input('start ?')
            b = driver.find_element_by_xpath('//*[@id="company_web_top"]/div[2]/div[3]/div[1]/h1')

        print(b.text)
        f.write(str(b.text) + '||||')

        try:
            a = driver.find_element_by_xpath('//*[@id="company_web_top"]/div[2]/div[3]/div[3]/div[2]/div[2]/div/div')
            print(a.text)
            f.write(str(a.text) + '\n')
        except:
            print('暂无信息')
            f.write('暂无信息\n')

        print(str(i) + ' / ' + str(sheet1.nrows - 1))

saved_num = 0
with open('data2_get.txt', 'a+') as f:
    pass

with open('data2_get.txt', 'r') as f:
    saved_num = len(f.readlines())

driver.get('https://www.tianyancha.com/')
wait = input('start ?')

for i in range(saved_num + 1, sheet2.nrows):

    company_name = str(int(sheet2.cell(i, 0).value))
    if company_name[0] == '6' and len(company_name) == 6:
        continue

    with open('data2_get.txt', 'a+') as f:
        company_name = sheet2.cell(i, 1).value
        driver.get('https://www.tianyancha.com/search?key=' + company_name)
        try:
            driver.find_elements_by_xpath('//a[@tyc-event-ch="CompanySearch.Company"]')[0].click()
            print('open new tab')
        except:
            print('no company for select')
            wait = input('start ?')
            try:
                driver.find_elements_by_xpath('//a[@tyc-event-ch="CompanySearch.Company"]')[0].click()
            except:
                print('no company for select checked')
                f.write('暂无信息||||暂无信息\n')
                continue

        driver.close()

        driver.switch_to.window(driver.window_handles[0])

        try:
            b = driver.find_element_by_xpath('//*[@id="company_web_top"]/div[2]/div[3]/div[1]/h1')
        except:
            print('company name not found')
            wait = input('start ?')
            b = driver.find_element_by_xpath('//*[@id="company_web_top"]/div[2]/div[3]/div[1]/h1')

        print(b.text)
        f.write(str(b.text) + '||||')

        try:
            a = driver.find_element_by_xpath('//*[@id="company_web_top"]/div[2]/div[3]/div[3]/div[2]/div[2]/div/div')
            print(a.text)
            f.write(str(a.text) + '\n')
        except:
            print('暂无信息')
            f.write('暂无信息\n')

        print(str(i) + ' / ' + str(sheet2.nrows - 1))

# # ========Write xls========
#
# workbook = xlwt.Workbook()
#
# worksheet = workbook.add_sheet('test_out')
#
# for i in range(sheet1.nrows):
#     for j in range(sheet1.ncols):
#         worksheet.write(i, j, sheet1.cell(i, j).value)
#
# with open('data_get.txt', 'r') as f:
#     lines = f.readlines()
#
#     for i in range(1, sheet1.nrows):
#         for line in lines:
#             if lines[i-1].split('||||')[0] in sheet1.cell(i, 1).value:
#                 worksheet.write(i, 7, lines[i-1].split('||||')[1])
#                 break
#
# workbook.save('test_out.xls')