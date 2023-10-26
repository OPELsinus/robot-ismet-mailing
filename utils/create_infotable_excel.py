import os

import numpy
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment

from config import main_folder


def create_infotable_excel(excel_file: str):

    book = load_workbook(os.path.join(main_folder, excel_file))

    sheet = book.active

    suppliers, count = dict(), 0

    for row in range(2, sheet.max_row):
        # print(sheet[f'C{row}'].value)
        if sheet[f'D{row}'].value == 'Отклонён':
            if sheet[f'C{row}'].value in suppliers:
                suppliers.update({sheet[f'C{row}'].value: suppliers.get(sheet[f'C{row}'].value) + 1})
            else:
                suppliers.update({sheet[f'C{row}'].value: 1})

        if sheet[f'C{row}'].value is None:
            break

    print(suppliers)

    book = Workbook()
    sheet = book.active

    for row, vals in enumerate(suppliers.items()):

        sheet[f'A{row + 2}'].value = vals[0]
        sheet[f'B{row + 2}'].value = vals[1]

    sheet[f'A1'].value = 'Поставщики'
    sheet[f'B1'].value = 'Отклонён'
    sheet[f'A1'].font = Font(bold=True)
    sheet[f'B1'].font = Font(bold=True)

    sheet[f'A{len(suppliers) + 2}'].value = 'Общий итог'
    sheet[f'B{len(suppliers) + 2}'].value = sum([i for i in suppliers.values()])
    sheet[f'A{len(suppliers) + 2}'].font = Font(bold=True)
    sheet[f'B{len(suppliers) + 2}'].font = Font(bold=True)

    fill_color = PatternFill(start_color="00CCCCFF", end_color="00CCCCFF", fill_type="solid")

    sheet.column_dimensions['A'].width = max([len(i) for i in suppliers.keys()]) * 1.15
    sheet.column_dimensions['B'].width = 15
    sheet['B1'].alignment = Alignment(horizontal='center')

    sheet['A1'].fill = fill_color
    sheet['B1'].fill = fill_color
    sheet[f'A{len(suppliers) + 2}'].fill = fill_color
    sheet[f'B{len(suppliers) + 2}'].fill = fill_color

    book.save('sfsd.xlsx')


