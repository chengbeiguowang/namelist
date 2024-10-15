# This is a sample Python script.
import os.path
import sys

from dexnamelist import scdex
from dexscoretable import compute_dex_score

from lawnamelist import handle_law_namelist
from namelist import handle_namelist

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
current_directory = os.path.dirname(os.path.abspath(__file__))

race_name_array = ["数据交易", "数据要素", "数字城市", "信创+"]
xlsx_name_array = [
    "第四届数字四川创新大赛初赛专家评分表（数据交易赛道）.xlsx",
    "第四届数字四川创新大赛初赛专家评分表（数据要素赛道）.xlsx",
    "第四届数字四川创新大赛初赛专家评分表（数字城市赛道）.xlsx",
    "第四届数字四川创新大赛初赛专家评分表（信创+赛道）.xlsx"
]

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #scdex('E:\\tele\\参赛清单汇总整理- 0823@1552.xlsx', 'E:\要素大赛')
    compute_dex_score('E:\评分表')


'''
    # file_base = os.path.realpath(os.path.dirname(sys.argv[0])) + "\\"
    file_base = ".\\"
    print(file_base)
    handle_law_namelist('G:\数字政法赛道评分资料', 'G:\数字政法赛道评分资料\数字政法赛道专家评分表(1129).xlsx')

    print("done")
'''

'''
    for i in range(0, len(race_name_array)):
        race_name = race_name_array[i]
        xlsx_name = xlsx_name_array[i]
        if os.path.exists(file_base + race_name) and os.path.exists(file_base + xlsx_name):
            handle_namelist(file_base + race_name, file_base + xlsx_name)
'''

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
