import os.path

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.styles import Font

from dexnamelist import Team
from namelist import clean_up
from utils import fill_cell

BACK_SLASH = "\\"

race_array = ['工业制造',
              '现代农业',
              '商贸流通',
              '交通运输',
              '金融服务',
              '科技创新',
              '文化旅游',
              '医疗健康',
              '应急管理',
              '气象服务',
              '城市治理',
              '绿色低碳']


class DexTeam(Team):
    def __init__(self, name, race_name, race_partition, work_name):
        super.__init__(name, race_name, race_partition, work_name)
        self.score = []

    def __str__(self):
        return str(self.name) + ' ' + str(self.race_name) + ' ' + str(self.race_partition) + ' ' + str(self.work_name)


def compute_dex_score(root_dir):
    summary_table = root_dir + BACK_SLASH + '汇总表.xlsx'
    for race in race_array:
        score_table_base = root_dir + BACK_SLASH + race
        if os.path.isdir(score_table_base):
            compute_race_score(race, score_table_base, summary_table)


def compute_race_score(race_name, race_dir, summary_table):
    count = 0
    total_dict = dict()
    for score_table in os.listdir(race_dir):
        table_dict = read_score_table(race_dir + BACK_SLASH + str(score_table))
        total_dict.update(table_dict)
        count += 1

    if count != 5:
        print("error in score table num:" + str(count) + " path:" + race_dir)

    write_score_to_summary_table(total_dict, summary_table, race_name)


def write_score_to_summary_table(total_dict, summary_table, race_name):
    wb = openpyxl.load_workbook(summary_table)
    sheet = wb[race_name]
    clean_up(sheet, 1, 12)

    ### write header
    fill_cell(sheet.cell(1, column=1), '序号')
    fill_cell(sheet.cell(1, column=2), '赛道名称')
    fill_cell(sheet.cell(1, column=3), '赛题名称')
    fill_cell(sheet.cell(1, column=4), '团队名称')
    fill_cell(sheet.cell(1, column=5), '团队作品')

    column = 6
    size = 0

    for expert_name in total_dict:
        # 写入专家姓名
        fill_cell(sheet.cell(1, column=column), expert_name)
        score_array = total_dict[expert_name]
        row = 2
        index = 0
        size = len(score_array)
        for item in score_array:
            fill_cell(sheet.cell(row + index, column=1), item['num'])
            fill_cell(sheet.cell(row + index, column=2), item['race_name'])
            fill_cell(sheet.cell(row + index, column=3), item['race_partition'])
            fill_cell(sheet.cell(row + index, column=4), item['team_name'])
            fill_cell(sheet.cell(row + index, column=5), item['work_name'])
            fill_cell(sheet.cell(row + index, column=column), float(item['score']))
            index += 1

        column += 1

    fill_cell(sheet.cell(1, column=11), '平均分')

    for i in range(2, 1 + size + 1):
        sum_value = float(0)
        for j in range(6, 11):
            sum_value += float(sheet.cell(i, j).value)

        avg = sum_value / 5
        fill_cell(sheet.cell(i, column=11), str(avg))

    wb.save(summary_table)


def read_score_table(score_table_path) -> dict:
    wb = openpyxl.load_workbook(score_table_path, data_only=True)
    sheet = wb['Sheet1']

    table_dict = {}
    count = 0

    expert_name = sheet.cell(2, column=1).value.split(':')[1]

    team_array = []

    row = 5
    cell = sheet.cell(row, column=1)
    while cell.value is not None:
        # 序号
        num = sheet.cell(row, column=1).value
        # 赛道名称
        race_name = sheet.cell(row, column=2).value
        # 赛题名称
        race_partition = sheet.cell(row, column=3).value
        # 团队名称
        team_name = sheet.cell(row, column=4).value
        # 团队作品
        work_name = sheet.cell(row, column=5).value
        # 评分合计
        score = str(sheet.cell(row, column=10).value)

        team_item = dict()
        team_item['num'] = num
        team_item['race_name'] = race_name
        team_item['race_partition'] = race_partition
        team_item['team_name'] = team_name
        team_item['work_name'] = work_name
        team_item['score'] = score
        if not score.isdigit():
            print("not number:" + str(team_item))

        team_array.append(team_item)
        count += 1

        row += 1
        cell = sheet.cell(row, column=1)

    table_dict[expert_name] = team_array
    print("专家：" + str(expert_name) + " item count:" + str(count))

    return table_dict
