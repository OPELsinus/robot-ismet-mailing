import datetime
import os

from openpyxl import load_workbook

from config import main_folder


















































def get_last_excel():

    nearest_day = 10000

    excel_file = None

    for file in os.listdir(main_folder):

        if nearest_day < (datetime.datetime.now() - datetime.datetime.fromtimestamp(os.path.getctime(os.path.join(main_folder, file)))).days:
            nearest_day = (datetime.datetime.now() - datetime.datetime.fromtimestamp(os.path.getctime(os.path.join(main_folder, file)))).days
            excel_file = file

    return excel_file


def is_file_corrupted(excel_file: str):

    return True if os.path.getsize(os.path.join(main_folder, excel_file)) / 1024 <= 1 else False


def is_file_empty(excel_file: str):

    book = load_workbook(os.path.join(main_folder, excel_file))

    sheet = book.active

    return True if sheet['A2'].value is None else False



