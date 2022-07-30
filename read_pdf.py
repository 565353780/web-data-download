from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
import logging
import os

# fw = open('out1.txt', 'w', encoding='utf-8')
#
# with open('out.txt', 'r', encoding='utf-8') as f:
#     lines = f.readlines()
#
#     for line in lines:
#         fw.write(line.split('\n')[0])
#
# fw.close()


logging.propagate = False
logging.getLogger().setLevel(logging.ERROR)

file_list = os.listdir('./')
pdf_filename = None
txt_filename = None

pdf_num = 0

solved_num = 0

for file in file_list:
    if '.pdf' in file or '.PDF' in file:
        pdf_num += 1

for file in file_list:
    if '.pdf' in file:
        pdf_filename = file
        txt_filename = file.split('.pdf')[0] + '.txt'
        if txt_filename in file_list:
            solved_num += 1
            print(file + '  done!')
            print(str(solved_num) + ' / ' + str(pdf_num))
            continue
    elif '.PDF' in file:
        pdf_filename = file
        txt_filename = file.split('.PDF')[0] + '.txt'
        if txt_filename in file_list:
            solved_num += 1
            print(file + '  done!')
            print(str(solved_num) + ' / ' + str(pdf_num))
            continue
    else:
        continue

    solved_num += 1

    device = PDFPageAggregator(PDFResourceManager(), laparams=LAParams())
    interpreter = PDFPageInterpreter(PDFResourceManager(), device)

    doc = PDFDocument()
    parser = PDFParser(open(pdf_filename, 'rb'))
    parser.set_document(doc)
    try:
        doc.set_parser(parser)
    except:
        continue
    doc.initialize()

    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        with open(txt_filename, 'w', encoding="utf-8") as fw:
            print("num page:{}".format(len(list(doc.get_pages()))))
            for page in doc.get_pages():
                interpreter.process_page(page)

                layout = device.get_result()

                for x in layout:
                    if isinstance(x, LTTextBoxHorizontal):
                        results = x.get_text()
                        fw.write(results)

    print(file + '  done!')
    print(str(solved_num) + ' / ' + str(pdf_num))
