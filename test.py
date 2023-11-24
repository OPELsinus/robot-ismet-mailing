import calendar
import datetime
import os
import shutil
import time
import uuid
from pathlib import Path

import openpyxl
import pandas as pd
import psycopg2
from openpyxl.styles import Border, Side

from config import logger, global_path

# def table_create():
#     """
#         Just simple table creation if not exists
#         """
#     conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
#     table_create_query = f'''CREATE TABLE IF NOT EXISTS ROBOT.{robot_name.replace("-", "_")} (
#         id text PRIMARY KEY,
#         status text,
#         retry_count INTEGER,
#         error_message text,
#         comments text,
#         execution_time text,
#         finish_date text,
#         date_created text,
#         executor_name text,
#
#         branch text,
#         store text,
#         court text,
#         branch_1c text,
#         branch_sprut text,
#         store_type text,
#         limit_sum text,
#         search_date text,
#         sprut_sum text,
#         odines_sum text,
#         outcome text,
#         excel_file text
#         )
#
#          '''
#     c = conn.cursor()
#     c.execute(table_create_query)
#     conn.commit()
#     c.close()
#     conn.close()


str_path_mapping_excel_file = "\\\\vault.magnum.local\\Common\\Stuff\\_06_Бухгалтерия\\Для робота\\Процесс Сверка Лимитов\\Маппинг для лимита касс.xlsx"
common_network_folder = "\\\\vault.magnum.local\\Common\\Stuff\\_06_Бухгалтерия\\Для робота\\"
main_directory_folder = "\\\\vault.magnum.local\\Common\\Stuff\\_06_Бухгалтерия\\Для робота\\Процесс Сверка Лимитов\\"

local_main_directory_folder = global_path.joinpath(f'.agent\\robot-sverka-limitov\\Output')
number_of_days = 1
months = ['', 'январь', 'февраль', 'март', 'апрель', 'май',
          'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']
thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))


def dispatch():
    # logger.info("Starting dispatcher")
    # table_create()
    search_date: str = (datetime.datetime.now() - datetime.timedelta(days=number_of_days)).strftime("%d.%m.%Y")
    # *  Find the corresponding main file, and if not exists, create new from mapping file.
    current_month: int = datetime.datetime.now().month
    current_month_name = months[int(search_date.split('.')[1])]
    current_year: int = int(search_date.split('.')[2])
    mapping_df = pd.read_excel(str_path_mapping_excel_file, sheet_name="Склад")
    main_working_file = None

    files = os.listdir(local_main_directory_folder)
    for item in files:
        if current_month_name in str(item).lower():

            if "~$" in item:
                item = item.replace("~$", "")

            main_working_file = os.path.join(local_main_directory_folder, item)

            break

    print(main_working_file)

    str_now = datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S')
    print(str_now)
    # exit()
    if not main_working_file:
        # * If there is no file related to current month
        file_name = f"лимит ГК филиалов {current_month_name.capitalize()} {current_year}.xlsx"
        main_working_file = os.path.join(local_main_directory_folder,
                                         file_name)
        logger.info(f"Главный файл не найден, создаем новый: {main_working_file}")
        # main_df = mapping_df.copy(deep=True)
        # main_df.drop(columns=['Склад', 'Площадка', 'Компания(Спрут)', 'Организация(1с)'], inplace=True)
        number_days_current_month = calendar.monthrange(year=current_year, month=current_month)[1]
        # for i in range(1, number_days_current_month + 1):
        #     new_column = datetime.datetime(day=i, month=current_month, year=current_year).strftime("%d.%m.%Y")
        #     main_df[new_column] = None
        # main_df.to_excel(main_working_file, sheet_name="Sheet1", index=False)
        time.sleep(3)
        # wb = openpyxl.load_workbook(main_working_file)
        # ws = wb["Sheet1"]
        # for idx, column in enumerate(ws.columns):
        #     column_letter = column[0].column_letter
        #     if idx > 3:
        #         ws.column_dimensions[column_letter].width = 15
        #     else:
        #         ws.column_dimensions[column_letter].width = 10
        #     for cell in column:
        #         cell.border = thin_border
        #
        # wb.save(main_working_file)
        # shared_full_path = str(Path(main_directory_folder).joinpath(file_name))

        # shutil.copyfile(main_working_file, shared_full_path)

        logger.info("Создан новый файл")
    else:
        logger.info("Найден главный файл")

    # * Need to create transactions based on mapping data
    # * Establish connection
    # conn = psycopg2.connect(host=db_host, port=db_port, database=db_name, user=db_user, password=db_pass)
    # c = conn.cursor()
    str_now = datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S')
    # transaction_count: int = 0
    for i, row in mapping_df.iterrows():

        branch = str(row[1])
        store = str(row[2])
        court = str(row[3])
        branch_sprut = str(row[5])
        branch_1c = str(row[4])
        store_type = str(row[6])
        limit = str(row[7])

        print(branch)

        # * need to check whether we already have a tr in db
        # find_query = f"Select id from ROBOT.{robot_name.replace('-', '_')} where search_date = '{search_date}' AND branch = '{branch}'"
        #
        # c.execute(find_query)
        # result = c.fetchone()
        # if result is None:
        #     # * insert a transaction to db
        insert_query = f"""Insert into ROBOT.robot_sverka_limitov (id, branch, store, court, branch_1c, branch_sprut, store_type, limit_sum, search_date, status, retry_count, date_created, excel_file) values ('{uuid.uuid4()}', '{branch}', '{store}', '{court}',
          '{branch_1c}','{branch_sprut}', '{store_type}', '{limit}', '{search_date}' , 'New', 0, '{str_now}', '{main_working_file}') """
        print(insert_query)
        #     c.execute(insert_query)
        #     conn.commit()
        #     transaction_count += 1

    # # * close db connection
    # c.close()
    # conn.close()
    # logger.info(f"Added {transaction_count} rows to  the DB")
    logger.info("Dispatcher ended")


if __name__ == '__main__':

    dispatch()


