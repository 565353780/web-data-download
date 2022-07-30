import os
import pdfplumber

# ========change to one-line-txt and cut from '在其他主体中的权益 '========

# if not os.path.exists('./changed_txt/'):
#     os.makedirs('./changed_txt/')
#
# txt_list = os.listdir('./')
#
# total_num = 0
#
# true_txt_num = 0
#
# not_used_txt_list = []
#
# for txt in txt_list:
#     if '.txt' in txt:
#         total_num += 1
#
# current_read_num = 0
#
# for txt in txt_list:
#     if '.txt' in txt:
#         with open(txt, 'r', encoding='utf-8') as f:
#             lines = f.readlines()
#
#             start_save = False
#
#             with open('./changed_txt/' + txt, 'w', encoding='utf-8') as fw:
#
#                 for line in lines:
#                     if '在其他主体中的权益 ' in line:
#
#                         if not start_save:
#                             true_txt_num += 1
#
#                         start_save = True
#
#                     if start_save:
#                         fw.write(line.split('\n')[0])
#
#             if not start_save:
#                 not_used_txt_list.append(txt)
#
#                 os.remove('./changed_txt/' + txt)
#
#         current_read_num += 1
#
#         print('\r' + str(current_read_num) + ' / ' + str(total_num), end='')
#
# print('')
#
# print('true_txt_num : ' + str(true_txt_num))
#
# with open('./not_used_txt_list.txt', 'w') as f:
#     for not_used_txt in not_used_txt_list:
#         f.write(not_used_txt + '\n')
#
# print('not used txt num : ' + str(len(not_used_txt_list)))

# ========pdfplumber========

# with pdfplumber.open(os.getcwd() + '/600006_2018_n.pdf') as pdf:
#     page=pdf.pages[138] #提取pdf第17页中的表格
#     for row in page.extract_tables():
#         print(row)

load_all_charts = True

txt_name = 'company_msgs_saved.txt'

if load_all_charts:
    txt_name = 'full_company_msgs_saved.txt'

file_list = os.listdir('./')

pdf_list = []

for file in file_list:
    if '.pdf' in file or '.PDF' in file:
        pdf_list.append(file)

total_pdf_num = len(pdf_list)

solved_pdf_id = 0

invalid_tag = [None, '']

if not os.path.exists(txt_name):
    with open(txt_name, 'a+') as f:
        pass

with open(txt_name, 'r') as f:
    lines = f.readlines()

    if len(lines) > 0:
        solved_pdf_id = int(lines[len(lines) - 1].split('|FuLi|')[0])

for i in range(solved_pdf_id, len(pdf_list)):
    path = pdf_list[i]

    if '_2018_n.pdf' not in path:
        continue

    pdf = pdfplumber.open(path)

    total_page_num = len(pdf.pages)

    for j in range(total_page_num - 1, -1, -1):

        pdf_tables = pdf.pages[j].extract_tables()

        for pdf_table in pdf_tables:

            if len(pdf_table[0]) != 7:
                continue

            if load_all_charts:

                for k in range(len(pdf_table)):

                    if pdf_table[k][0] not in invalid_tag and pdf_table[k][1] not in invalid_tag:

                        with open(txt_name, 'a+') as f:

                            write_msg = ''

                            write_msg += str(i)

                            for l in range(7):
                                try:
                                    write_msg += '|FuLi|' + pdf_table[k][l].replace(' ', '').replace('\n', '')
                                except:
                                    write_msg += '|FuLi|'

                            write_msg += '\n'

                            try:
                                f.write(write_msg)
                            except:
                                continue

            else:

                valid_title_line_id = 0

                while pdf_table[valid_title_line_id][0] in invalid_tag or pdf_table[valid_title_line_id][1] in invalid_tag:
                    valid_title_line_id += 1

                    if valid_title_line_id > len(pdf_table) - 1:
                        break

                if valid_title_line_id > len(pdf_table) - 1:
                    continue

                pdf_table_copy = pdf_table

                if valid_title_line_id > 0:
                    pdf_table_copy = []

                    valid_title_line_id = 0

                    while valid_title_line_id < len(pdf_table):

                        current_merged_line = ['', '', '', '', '', '', '']

                        is_invalid_line = True

                        for k in range(7):
                            if pdf_table[valid_title_line_id][k] not in invalid_tag:
                                is_invalid_line = False

                                break

                        while not is_invalid_line:

                            for k in range(7):
                                if pdf_table[valid_title_line_id][k] != None:
                                    current_merged_line[k] += pdf_table[valid_title_line_id][k]

                            valid_title_line_id += 1

                            if valid_title_line_id > len(pdf_table) - 1:
                                break

                            is_invalid_line = True

                            for k in range(7):
                                if pdf_table[valid_title_line_id][k] not in invalid_tag:
                                    is_invalid_line = False

                                    break

                        valid_title_line_id += 1

                        pdf_table_copy.append(current_merged_line)

                first_msg = pdf_table_copy[0][0].replace(' ', '').replace('\n', '')
                second_msg = pdf_table_copy[0][1].replace(' ', '').replace('\n', '')

                if first_msg == '子公司名称' and second_msg == '主要经营地':

                    for k in range(1, len(pdf_table_copy)):

                        if pdf_table_copy[k][0] not in invalid_tag:

                            with open(txt_name, 'a+') as f:

                                write_msg = ''

                                write_msg += str(i)

                                for l in range(7):
                                    write_msg += '|FuLi|' + pdf_table_copy[k][l].replace(' ', '').replace('\n', '')

                                write_msg += '\n'

                                try:
                                    f.write(write_msg)
                                except:
                                    continue

        current_text = pdf.pages[j].extract_text()

        if current_text == None:
            continue

        if '在其他主体中的权益 ' in current_text:
            break

        print('\r' + str(i + 1) + ' / ' + str(total_pdf_num) + '\t' + str(total_page_num - j) + ' / ' + str(total_page_num), end='')

    pdf.close()