import xlrd
import xlwt
import os

# ========Read xls========

book1 = xlrd.open_workbook('test1.xls')
sheet1 = book1.sheets()[0]

# book2 = xlrd.open_workbook('test2.xls')
# sheet2 = book2.sheets()[0]

# ========Write xls========

workbook1 = xlwt.Workbook()

worksheet1 = workbook1.add_sheet('test1_out')

for i in range(sheet1.nrows):
    for j in range(sheet1.ncols):
        worksheet1.write(i, j, sheet1.cell(i, j).value)

matched_num = 0

with open('data1_get.txt', 'r') as f:
    lines = f.readlines()

    for i in range(1, sheet1.nrows):

        current_exist_line_num = 0

        for line in lines:
            line_cut = line.split('||||')

            if sheet1.cell(i, 1).value == line_cut[0]:

                current_exist_line_num += 1

                if current_exist_line_num == 1:
                    matched_num += 1
                
                sheng_cut = ''
                other_cut = line_cut[1]
                if '省' in line_cut[1]:
                    sheng_cut = other_cut.split('省')[0] + '省'
                    other_cut = other_cut.split('省')[1]

                    worksheet1.write(i, 7 + current_exist_line_num, sheng_cut)
                current_exist_line_num += 1
                
                shi_cut = ''
                if '市' in other_cut:
                    shi_cut = other_cut.split('市')[0] + '市'
                    other_cut = other_cut.split('市')[1]

                    worksheet1.write(i, 7 + current_exist_line_num, shi_cut)
                current_exist_line_num += 1
                
                qu_cut = ''
                if '区' in other_cut:
                    qu_cut = other_cut.split('区')[0] + '区'
                    other_cut = other_cut.split('区')[1]

                    worksheet1.write(i, 7 + current_exist_line_num, qu_cut)
                current_exist_line_num += 1

                worksheet1.write(i, 7 + current_exist_line_num, line_cut[1])
                break

        print('\r' + str(i + 1) + ' / ' + str(sheet1.nrows) + '\t' + str(matched_num), end='')

print('')

matched_num = 0

with open('full_company_msgs_saved.txt', 'r') as f:
    lines = f.readlines()

    for i in range(1, sheet1.nrows):

        current_exist_line_num = 0

        for line in lines:
            line_cut = line.split('|FuLi|')

            if sheet1.cell(i, 1).value == line_cut[1]:

                current_exist_line_num += 1

                if current_exist_line_num == 1:
                    matched_num += 1

                worksheet1.write(i, 12 + current_exist_line_num, line_cut[2])

        print('\r' + str(i + 1) + ' / ' + str(sheet1.nrows) + '\t' + str(matched_num), end='')

print('')

workbook1.save('test1_out.xls')

# workbook2 = xlwt.Workbook()

# worksheet2 = workbook2.add_sheet('test2_out')

# for i in range(sheet2.nrows):
#     for j in range(sheet2.ncols):
#         worksheet2.write(i, j, sheet2.cell(i, j).value)

# matched_num = 0

# with open('full_company_msgs_saved.txt', 'r') as f:
#     lines = f.readlines()

#     for i in range(1, sheet2.nrows):

#         current_exist_line_num = 0

#         for line in lines:
#             line_cut = line.split('|FuLi|')

#             if line_cut[2] == '':
#                 continue

#             if sheet2.cell(i, 1).value == line_cut[1]:

#                 current_exist_line_num += 1

#                 if current_exist_line_num == 1:
#                     matched_num += 1

#                 worksheet2.write(i, 6 + current_exist_line_num, line_cut[2])

#         print('\r' + str(i + 1) + ' / ' + str(sheet2.nrows) + '\t' + str(matched_num), end='')

# print('')

# workbook2.save('test2_out.xls')