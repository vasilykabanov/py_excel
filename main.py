# This is a sample Python script.
import os
import pandas
import fnmatch
# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.


def read_excel():
    excel_file = input("Введите название файла Excel: ")
    # table.xlsx
    excel_sheet = input("Введите название столбца откуда нужно считать ячейки: ")
    # A1
    path = input("Введите полный путь к папке с файлами: ")
    # /Users/vvkaban2/Desktop/fold

    excel_data_df = pandas.read_excel(excel_file)
    excel_data_df.head()
    excel_list = excel_data_df[excel_sheet].tolist()
    # print(excel_list)

    result = []
    log = []
    more_than_two = []
    for root, dirs, files in os.walk(path):
        for x in excel_list:
            count = 0
            for name in files:
                if fnmatch.fnmatch(name, '*.log'):
                    if name in log:
                        continue
                    else:
                        log.append(os.path.join(name))
                        continue
                if fnmatch.fnmatch(name, '*' + str(x) + '*'):
                    result.append(os.path.join(name))
                    count += 1
                    if count > 2:
                        styled = (excel_data_df.style.
                                  applymap(lambda i: 'background-color: %s' % 'green' if i == x else ''))
                        more_than_two.append(name)
                        styled.to_excel('table.xlsx', engine='openpyxl')

    f = open('logfiles.txt', 'tw', encoding='utf-8')
    for item in log:
        f.write("%s\n" % item)
    f.close()

    print('Все файлы с расширением ".log": ' + str(log))
    print('Файлы, которые встречались больше двух раз: ' + str(more_than_two))
    return result


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    read_excel()
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
