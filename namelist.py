import os

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.worksheet.worksheet import Worksheet


def handle_namelist(race_path, workbook_path):
    # sheet
    wb = openpyxl.load_workbook(workbook_path)
    sheet = wb['Sheet1']

    # 清理表格
    clean_up(sheet, 6, 3)

    #
    row = 6
    for dir_name in os.listdir(race_path):
        res = dir_name.split(" ")
        # 序号
        fill_cell(sheet.cell(row, 1), res[0])
        # 团队名
        fill_cell(sheet.cell(row, 2), res[1])

        team_dir = race_path + "\\" + dir_name
        files = os.listdir(team_dir)
        cell = sheet.cell(row, column=3)

        if len(files) == 1:
            fill_cell_link(cell, files[0], team_dir + "\\" + files[0])
        else:
            zip_file = None
            for file in files:
                if os.path.splitext(file)[-1][1:] == "zip":
                    zip_file = file
                    break

            if zip_file is None:
                fill_cell_link(cell, zip_file, team_dir)

            else:
                fill_cell_link(cell, files[0], team_dir)

        row += 1

    wb.save(workbook_path)

def fill_cell_link(cell, file_name, link):
    cell.value = file_name
    cell.hyperlink = link
    cell.style = "Hyperlink"
    cell.alignment = Alignment(horizontal='left', vertical='center',wrapText=True)

def fill_cell(cell, content):
    cell.value =content
    cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)

def alignment():
    return Alignment(horizontal='left', vertical='center', wrapText=True)

def is_empty(value) -> bool:
    return value is None or value == ''


def clean_up(sheet: Worksheet, row_start, column_max):
    _row_start = row_start
    cell = sheet.cell(_row_start, 1)
    while not is_empty(cell.value):
        for i in range(1, column_max + 1):
            cell = sheet.cell(_row_start, i)
            cell.value = None
        _row_start += 1
        cell = sheet.cell(_row_start, 1)
