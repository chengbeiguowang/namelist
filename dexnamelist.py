import os

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font

from utils import fill_cell, clean_up

FONT_SIZE = 16
COLUMN_SUM = 12


class Team:
    def __init__(self, name, race_name, race_partition, work_name):
        self.name = name
        self.race_name = race_name
        self.race_partition = race_partition
        self.work_name = work_name

    def __str__(self):
        return str(self.name) + ' ' + str(self.race_name) + ' ' + str(self.race_partition) + ' ' + str(self.work_name)

def scdex():
    race_dict = read_team_output('', 'E:\要素大赛\科技创新\参赛清单汇总整理- 0821@1230.xlsx')
    write_to_score_sheet(race_dict, 'E:\要素大赛\科技创新\数据要素x大赛初赛专家评分表.xlsx')

def write_to_score_sheet(race_dict, workbook_path):
    wb = openpyxl.load_workbook(workbook_path)
    sheet = wb['Sheet1']

    clean_up(sheet, 5, 10)

    race_partition_dict = race_dict['科技创新']
    row = 5
    for key in race_partition_dict:
        partition_array = race_partition_dict[key]
        for i in range(len(partition_array)):
            team: Team = partition_array[i]
            if team.work_name is None or team.work_name == '':
                continue
            fill_cell(sheet.cell(row, column=1), str(row - 4))
            fill_cell(sheet.cell(row, column=2), str(team.race_name))
            fill_cell(sheet.cell(row, column=3), str(team.race_partition))
            fill_cell(sheet.cell(row, column=4), str(team.name))
            fill_cell(sheet.cell(row, column=5), str(team.work_name))

            # 累加和
            cell = sheet.cell(row, column=10)
            cell.value = "=SUM({}:{})".format("F" + str(row), "I" + str(row))
            cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
            font = Font('Times New Roman', color='000000', bold=False, size=FONT_SIZE)
            cell.font = font

            sheet.row_dimensions[row].height = 100
            row += 1

    wb.save(workbook_path)


def read_team_output(race_path, workbook_path) -> dict:
    # sheet
    wb = openpyxl.load_workbook(workbook_path)
    sheet = wb['参赛团队列表824']
    team_set = set()
    race_dict = {}

    for row in range(2, sheet.max_row):
        # 参赛团队
        team_name = sheet.cell(row, column=3).value
        if team_name in team_set:
            continue

        team_set.add(team_name)

        # 赛道名称
        race_name = sheet.cell(row, column=1).value
        # 参赛方向
        race_partition = sheet.cell(row, column=2).value

        # 作品名称
        work_name = sheet.cell(row, column=7).value

        team = Team(team_name, race_name, race_partition, work_name)

        if race_dict.get(race_name) is None:
            race_dict[race_name] = {}

        if not race_dict[race_name].__contains__(race_partition):
            race_dict[race_name][race_partition] = []

        race_dict[race_name][race_partition].append(team)

    print(str(race_dict))
    return race_dict
