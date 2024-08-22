from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet


def fill_cell(cell, content, font_size=16):
    cell.value = content
    cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
    font = Font('宋体', color='000000', bold=False, size=font_size)
    cell.font = font


def fill_cell_link(cell, file_name, link, font_size=16):
    cell.value = file_name
    cell.hyperlink = link
    font = Font('宋体', color='0563C1', bold=False, size=font_size)
    cell.font = font
    # cell.style = "Hyperlink"
    cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)


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
