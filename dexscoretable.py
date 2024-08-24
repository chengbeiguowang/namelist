import os.path

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.styles import Font

from dexnamelist import Team

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
    for race in race_array:
        score_table_base = root_dir + BACK_SLASH + race
        if os.path.isdir(score_table_base):
            compute_race_score(score_table_base)


def compute_race_score(race_dir):
    count = 0
    for score_table in os.listdir(race_dir):
        read_score_table(race_dir + BACK_SLASH + str(score_table))
        count += 1

    if count != 5:
        print("error in score table num:" + str(count) + " path:" + race_dir)


def read_score_table(score_table_path) -> dict:
    wb = openpyxl.load_workbook(score_table_path, data_only=True)
    sheet = wb['Sheet1']

    table_dict = {}
    count = 0

    expert_name = sheet.cell(2, column=1).value.split(':')[1]
    print(expert_name)

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

    print("专家：" + str(expert_name) + " item count:" + str(count))

    return table_dict
