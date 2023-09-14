# This is a sample Python script.
from namelist import handle_namelist


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    file_base = "E:\\pythonProject\\创新大赛\\"
    race_name = "数据交易"
    xlsx_name = "第四届数字四川创新大赛初赛专家评分表.xlsx"

    handle_namelist(file_base + race_name, file_base + xlsx_name)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
