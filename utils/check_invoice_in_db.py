import datetime

from config import logger, engine_kwargs, robot_name, global_env_data
from utils.check_if_excel_good import get_last_excel
from utils.create_infotable_excel import create_infotable_excel

from sqlalchemy import create_engine, Column, Integer, String, DateTime, MetaData, Table, Date, Boolean, select, BigInteger
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


def check_invoice_in_db():

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

    from_date = datetime.date(2023, 10, 23)
    condition = IsmetTable.DATE_INVOICE >= from_date

    select_query = session_ismet.query(IsmetTable).filter(condition).all()

    for a in select_query:
        print(a.ID_INVOICE, a.NUMBER_INVOICE, a.DATE_INVOICE, sep=' | ')


if __name__ == '__main__':

    check_invoice_in_db()

