import requests
import os

try:
    with open('data_get1.txt', 'r') as f:
        lines = f.readlines()

        new_lines = list(set(lines))

        with open('data_get1_filter.txt', 'w') as f1:
            for line in new_lines:
                f1.write(line)

    with open('data_get1_filter.txt', 'r') as f:
        lines = f.readlines()

        file_list = os.listdir('./')

        saved_num = len(file_list)

        for line in lines:
            file_list = os.listdir('./')

            url = line.split('\n')[0]

            file_name = url.split('/')[len(url.split('/')) - 1]

            if file_name not in file_list:

                f_down = requests.get(url)

                with open(file_name, 'wb') as fwrite:
                    fwrite.write(f_down.content)

                saved_num += 1

                print('\r' + str(saved_num) + ' / ' + str(len(lines)), end='')

        print('')
except:
    pass

with open('data_get2.txt', 'r') as f:
    lines = f.readlines()

    new_lines = list(set(lines))

    with open('data_get2_filter.txt', 'w') as f1:
        for line in new_lines:
            f1.write(line)

with open('data_get2_filter.txt', 'r') as f:
    lines = f.readlines()

    file_list = os.listdir('./')

    saved_num = len(file_list)

    for line in lines:
        file_list = os.listdir('./')

        url = line.split('\n')[0]

        file_name = url.split('/')[len(url.split('/')) - 1]

        if file_name not in file_list:
            f_down = requests.get(url)

            with open(file_name, 'wb') as fwrite:
                fwrite.write(f_down.content)

            saved_num += 1

            print('\r' + str(saved_num) + ' / ' + str(len(lines)), end='')

    print('')