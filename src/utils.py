from datetime import datetime
import datetime
import os
import win32com.client
import sys
import logging
import sqlalchemy as sa
from sqlalchemy import create_engine, text, Engine, Connection, Table
from pandas.io.sql import SQLTable
from sqlalchemy import DATE
import yaml
from urllib.parse import quote
import pandas as pd
from src.telegram_bot import *

proyect_dir = os.path.dirname(os.path.abspath(__file__))
today = datetime.date.today()
log_dir = os.path.join(proyect_dir, '..', 'log', 'logs_main.log')
yml_credentials_dir = os.path.join(
    proyect_dir, '..', 'config', 'credentials.yml')

sys.path.append(proyect_dir)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
subjectt_1 = 'Scheduled Report: Report_Total_Col_3'
subjectt_2 = 'Scheduled Report: Report_Total_Col_4'

logging.basicConfig(
    level=logging.INFO,
    filename=(log_dir),
    format="%(asctime)s - %(levelname)s -  %(message)s",
    datefmt='%d-%b-%y %H:%M:%S'
)


def get_engine(username: str, password: str, host: str, database: str, port: str = 3306, *_) -> Engine:
    return sa.create_engine(f"mysql+pymysql://{username}:{quote(password)}@{host}:{port}/{database}?autocommit=true")


with open(yml_credentials_dir, 'r') as f:

    try:
        config = yaml.safe_load(f)
        source1 = config['source1']
    except yaml.YAMLError as e:
        logging.error(str(e), exc_info=True)


def engine_1() -> Connection:
    return get_engine(**source1).connect()


def to_sql_replace(table: SQLTable, con: Engine | Connection, keys: list[str], data_iter):

    satable: Table = table.table
    ckeys = list(map(lambda s: s.replace(' ', '_'), keys))
    data = [dict(zip(ckeys, row)) for row in data_iter]
    values = ', '.join(f':{nm}' for nm in ckeys)
    stmt = f"REPLACE INTO {satable.name} VALUES ({values})"

    con.execute(text(stmt), data)


def saveattachemnts():

    for message in messages:

        if message.Subject == subjectt_1 or message.Subject == subjectt_2 and message.Senton.date() == today:
            attachments = message.Attachments
            attachment = attachments.Item(1)
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(
                    proyect_dir, '..', 'data', str(attachment)))
                if message.Subject == subjectt_1 or message.Subject == subjectt_2 and message.Unread:
                    message.Unread = False
                    message.Delete()
                break


def load():

    with engine_1() as con:

        asyncio.run(enviar_mensaje('Cargue hikvision'))
        file_name = os.listdir(os.path.join(proyect_dir, '..', 'data'))[0]
        df_hik = pd.read_csv(os.path.join(
            proyect_dir, '..', 'data', file_name))

        asyncio.run(enviar_mensaje(f'{len(df_hik)} datos leidos'))

        if len(df_hik) == 0:
            asyncio.run(enviar_mensaje('Reporte sin datos'))
            asyncio.run(enviar_mensaje(
                ("____________________________________")))
            os.remove(os.path.join(
                proyect_dir, '..', 'data', file_name))
            sys.exit()

        df_hik.columns = df_hik.columns.str.replace(' ', '_')

        df_hik['DATE_AND_HOUR'] = pd.to_datetime(
            df_hik['DATE_AND_HOUR'], format='%a, %d %b %Y %H:%M')
        load_date = df_hik['DATE_AND_HOUR'][0]

        df_hik['HOLDS'] = None

        df_hik = df_hik[['DATE_AND_HOUR', 'CALL_ID', 'ABANDONED', 'TIME_TO_ABANDON',
                         'SKILL', 'CALL_TYPE', 'CAMPAIGN', 'CALLS', 'AFTER_CALL_WORK_TIME',
                         'HOLDS', 'TALK_TIME', 'HANDLE_TIME', 'TRANSFERS', 'AGENT_EMAIL',
                         'QUEUE_WAIT_TIME', 'SPEED_OF_ANSWER', 'TOTAL_QUEUE_TIME',
                         'QUEUE_CALLBACK_WAIT_TIME', 'IVR_TIME', 'HOLD_TIME', 'AGENT_GROUP']]

        # df_hik = df_hik[df_hik['SKILL'] != 'Password Reset']

        load_date = load_date.strftime('%Y-%m-%d')

        count = pd.read_sql_query(
            f"SELECT COUNT(*) FROM bbdd_cos_bog_hikvision_bi.tb_cos_raw_data_enhanced where date(DATE_AND_HOUR) = '{load_date}';", con)['COUNT(*)'][0]

        con.execute(
            text(f"DELETE FROM bbdd_cos_bog_hikvision_bi.tb_cos_raw_data_enhanced WHERE DATE(DATE_AND_HOUR) = '{load_date}' limit {count}"))
        print(load_date)

    with engine_1() as con:

        df_hik.to_sql(name='tb_cos_raw_data_enhanced', con=con,
                      if_exists='append', index=False)

        count_ = pd.read_sql_query(
            f"SELECT COUNT(*) FROM bbdd_cos_bog_hikvision_bi.tb_cos_raw_data_enhanced where date(DATE_AND_HOUR) = '{load_date}';", con)['COUNT(*)'][0]

        asyncio.run(enviar_mensaje(
            f'{load_date} \n {len(df_hik)} datos cargados \n {count_} datos en tabla'))
        asyncio.run(enviar_mensaje(
            str("____________________________________")))

        os.remove(os.path.join(
            proyect_dir, '..', 'data', file_name))
