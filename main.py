import datetime

from config import logger, engine_kwargs, robot_name, smtp_host, smtp_author
from tools.smtp import smtp_send
from utils.check_if_excel_good import get_last_excel
from utils.check_invoice_in_db import check_invoice_in_db
from utils.create_infotable_excel import create_infotable_excel

from sqlalchemy import create_engine, Column, Integer, String, DateTime, MetaData, Table, Date, Boolean, select
from sqlalchemy.orm import declarative_base, sessionmaker

from utils.divide import divide_excel_by_suppliers
from utils.get_all_emails_sprut import get_all_emails

Base = declarative_base()


class Table(Base):
    __tablename__ = robot_name.replace('-', '_')

    date_created = Column(Date, default=None)
    id_invoice = Column(String(512), primary_key=True)
    reason_invoice = Column(String(512), default=None)
    store_name = Column(String(512), default=None)
    supplier_name = Column(String(512), default=None)

    status = Column(String(16), default=None)

    @property
    def dict(self):
        m = self.__dict__.copy()
        return m


if __name__ == '__main__':

    # get_last_excel()

    Session = sessionmaker()

    engine = create_engine(
        'postgresql+psycopg2://{username}:{password}@{host}:{port}/{base}'.format(**engine_kwargs),
        connect_args={'options': '-csearch_path=robot'}
    )
    Base.metadata.create_all(bind=engine)
    Session.configure(bind=engine)
    session = Session()

    excel_file, invoices = check_invoice_in_db()

    for key, val in invoices.items():
        print(key, val)
        session.add(Table(
            date_created=datetime.datetime.now(),
            id_invoice=key,
            reason_invoice=val[0],
            store_name=val[1],
            supplier_name=val[2],
            status='new'
        ))
    session.commit()

    suppliers_excels = divide_excel_by_suppliers(excel_file)

    create_infotable_excel(excel_file)

    emails = get_all_emails()

    print(emails)
    for key, emails_ in emails.items():
        print('----------')
        print(key)
        emls = []
        print(f"smtp_send('assdf', url=smtp_host, to={[email.strip() for email in emails_.split(',')]}, subject=f'Исмет Рассылка Тест', username=smtp_author)")

    # smtp_send('assdf', url=smtp_host, to=['Abdykarim.D@magnum.kz'], subject=f'Исмет Рассылка Тест', username=smtp_author)
