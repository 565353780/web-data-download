import xlrd
import xlwt

# ========Read xls========

book1 = xlrd.open_workbook('test1.xls')
sheet1 = book1.sheets()[0]

book2 = xlrd.open_workbook('test2.xls')
sheet2 = book2.sheets()[0]

# ========Write xls========

workbook1 = xlwt.Workbook()

worksheet1 = workbook1.add_sheet('test1_out')

for i in range(sheet1.nrows):
    for j in range(sheet1.ncols):
        worksheet1.write(i, j, sheet1.cell(i, j).value)

matched_num = 0

with open('company_msgs_saved.txt', 'r') as f:
    lines = f.readlines()

    for i in range(1, sheet1.nrows):

        current_exist_line_num = 0

        for line in lines:
            if line.split('|FuLi|')[2] == '':
                continue

            if line.split('|FuLi|')[1] in sheet1.cell(i, 1).value:

                current_exist_line_num += 1

                if current_exist_line_num == 1:
                    matched_num += 1

            worksheet1.write(i, 6 + current_exist_line_num, line.split('|FuLi|')[2])

        print('\r' + str(i + 1) + ' / ' + str(sheet1.nrows) + '\t' + str(matched_num), end='')

print('')

workbook1.save('test1_out.xls')

workbook2 = xlwt.Workbook()

worksheet2 = workbook2.add_sheet('test2_out')

for i in range(sheet2.nrows):
    for j in range(sheet2.ncols):
        worksheet2.write(i, j, sheet2.cell(i, j).value)

matched_num = 0

with open('company_msgs_saved.txt', 'r') as f:
    lines = f.readlines()

    for i in range(1, sheet2.nrows):
        for line in lines:
            if line.split('|FuLi|')[2] == '':
                continue

            if line.split('|FuLi|')[1] in sheet2.cell(i, 1).value:
                matched_num += 1

                worksheet2.write(i, 7, line.split('|FuLi|')[2])
                break

        print('\r' + str(i + 1) + ' / ' + str(sheet2.nrows) + '\t' + str(matched_num), end='')

print('')

workbook2.save('test2_out.xls')