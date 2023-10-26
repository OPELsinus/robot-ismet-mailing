import datetime

from config import logger, engine_kwargs, robot_name
from utils.check_if_excel_good import get_last_excel
from utils.create_infotable_excel import create_infotable_excel

from sqlalchemy import create_engine, Column, Integer, String, DateTime, MetaData, Table, Date, Boolean, select
from sqlalchemy.orm import declarative_base, sessionmaker


Base = declarative_base()


class Table(Base):
    __tablename__ = robot_name.replace('-', '_')

    file_path = Column(String(512), primary_key=True)
    date_created = Column(Date, default=None)

    id_invoice = Column(String(512), default=None)
    reason_invoice = Column(String(512), default=None)
    supplier_name = Column(String(512), default=None)

    status = Column(String(16), default=None)

    @property
    def dict(self):
        m = self.__dict__.copy()
        return m


if __name__ == '__main__':

    get_last_excel()

    # open_last_excel()

    # create_infotable_excel(r'C:\Users\Abdykarim.D\Documents\2023-09-07_ismet.xlsx')

    Session = sessionmaker()

    engine = create_engine(
        'postgresql+psycopg2://{username}:{password}@{host}:{port}/{base}'.format(**engine_kwargs),
        connect_args={'options': '-csearch_path=robot'}
    )
    Base.metadata.create_all(bind=engine)
    Session.configure(bind=engine)
    session = Session()

    # session.add(Table(
    #     date_created=datetime.datetime.now(),
    #     file_path='kek1.xlsx',
    #     id_invoice='KEKUS',
    #     reason_invoice='LOOL',
    #     supplier_name='DEALER'
    # ))
    # session.commit()

    select_query = session.query(Table).all()

    for a in select_query:
        print(a.file_path)


