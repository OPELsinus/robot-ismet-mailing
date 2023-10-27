import datetime
import os

from openpyxl import load_workbook, Workbook

from config import logger, engine_kwargs, robot_name, global_env_data, main_folder
from utils.check_if_excel_good import get_last_excel
from utils.create_infotable_excel import create_infotable_excel

from sqlalchemy import create_engine, Column, Integer, String, DateTime, MetaData, Table, Date, Boolean, select, BigInteger, distinct
from sqlalchemy.orm import declarative_base, sessionmaker


Base = declarative_base()


class IsmetTable(Base):

    __tablename__ = "parse_all"
    id = Column(Integer, primary_key=True)
    # add_time = Column(DateTime, default=None)
    edit_time = Column(DateTime, default=None)
    status = Column(String(128), default=None)
    error_message = Column(String(128), default=None)

    ID_INVOICE = Column(String(128))
    NUMBER_INVOICE = Column(String(128))
    URL_INVOICE = Column(String(128))
    C_NAME_SOURCE_INVOICE = Column(String(128))
    C_NAME_SHOP = Column(String(128))
    DATE_INVOICE = Column(DateTime)
    BAR_CODE_WARES = Column(BigInteger)
    NAME_WARES = Column(String(128))
    QUANTITY = Column(Integer)
    APPROVE_FLAG = Column(Integer, default=None)

    @property
    def dict(self):
        m = self.__dict__
        return m


def check_invoice_in_db(excel_file: str):

    """
        Creating connection to the ismet table
        Getting all rows from the excel with status 'Отклонён'
        Fetching the invoice number for that status from the ismet table
        Creating new excel with only rows with status 'Отклонён'
    """

    # * Creating connection to the ismet table

    Session_ismet = sessionmaker()

    engine_kwargs_ismet = {
        'username': global_env_data['postgre_db_username'],
        'password': global_env_data['postgre_db_password'],
        'host': global_env_data['postgre_ip'],
        'port': global_env_data['postgre_port'],
        'base': 'ismet'
    }

    engine = create_engine(
        'postgresql+psycopg2://{username}:{password}@{host}:{port}/{base}'.format(**engine_kwargs_ismet),
        connect_args={'options': '-csearch_path=public'}
    )

    Base.metadata.create_all(bind=engine)
    Session_ismet.configure(bind=engine)
    session_ismet = Session_ismet()

    # * Getting all rows with status 'Отклонён'

    book = load_workbook(os.path.join(main_folder, excel_file))
    sheet = book.active

    invoices = []

    for row in range(2, sheet.max_row):

        if sheet[f'A{row}'].value is None:
            break

        if sheet[f'D{row}'].value != 'Отклонён':
            continue

        # print(sheet[f'A{row}'].value)

        invoices.append(sheet[f'A{row}'].value)

    # * Fetching all nubmer invoices from the db

    from_date = datetime.date(2023, 10, 23)
    condition = IsmetTable.DATE_INVOICE >= from_date

    condition1 = IsmetTable.ID_INVOICE.in_(invoices)

    select_query = (
        session_ismet.query(IsmetTable)
        .filter(condition1)
        .filter(IsmetTable.NUMBER_INVOICE.isnot(None))
        .all()
    )
    invoices = dict()
    for ind, row in enumerate(select_query):
        if row.ID_INVOICE not in list(invoices.keys()):
            try:
                _ = int(row.NUMBER_INVOICE)
            except:
                invoices.update({row.ID_INVOICE: [row.NUMBER_INVOICE, row.C_NAME_SOURCE_INVOICE, row.DATE_INVOICE]})
            # print(a.ID_INVOICE, a.NUMBER_INVOICE, a.DATE_INVOICE, sep=' | ')

    # * Creating new Excel from rows with status 'Отклонён'

    rows_to_copy = []

    for row in range(2, sheet.max_row):

        if sheet[f'A{row}'].value is None:
            break

        if sheet[f'D{row}'].value != 'Отклонён':
            continue
        found = False
        for key, val in invoices.items():
            # print(key, val)
            if sheet[f'A{row}'].value == key:
                sheet[f'H{row}'].value = val[0]
                rows_to_copy.append(row)
                found = True
        if not found:
            rows_to_copy.append(row)

    print('deleting')
    print(rows_to_copy)
    book1 = Workbook()
    sheet1 = book1.active

    for ind, row in enumerate(rows_to_copy):
        for letter in 'ABCDEFGH':
            sheet1[f'{letter}{ind + 1}'].value = sheet[f'{letter}{row}'].value

    book1.save('dsfsdffsfd.xlsx')

    return invoices.keys()

# if __name__ == '__main__':
#
#     check_invoice_in_db()


