from pandas.io import sql
from sqlalchemy.orm import sessionmaker
from urllib.parse import quote_plus

import pandas as pd
import sqlalchemy as mysql
import shelve
import pyodbc
import copy
import os
import datetime
import logging


class SQLConnect:
    session = False
    engine = None
    conn = None
    cursor = None

    def __init__(self, conn_type, dsn=None):
        self.conn_type = conn_type

        if conn_type == 'alch':
            self.connstring = self.alchconnstr(
                '{SQL Server Native Client 11.0}', 1433, settings['Server'], settings['Database'], 'mssql'
                )
        elif conn_type == 'sql':
            self.connstring = self.sqlconnstr(settings['Server'], settings['Database'])
        elif conn_type == 'dsn':
            self.connstring = self.dsnconnstr(dsn)

    @staticmethod
    def alchconnstr(driver, port, server, database, flavor='mssql'):
        p = quote_plus(
                'DRIVER={};PORT={};SERVER={};DATABASE={};Trusted_Connection=yes;'
                .format(driver, port, server, database))

        return '{}+pyodbc:///?odbc_connect={}'.format(flavor, p)

    @staticmethod
    def sqlconnstr(server, database):
        return 'driver={0};server={1};database={2};autocommit=True;Trusted_Connection=yes'.format('{SQL Server}',
                                                                                                  server, database)

    @staticmethod
    def dsnconnstr(dsn):
        return 'DSN={};DATABASE=default;Trusted_Connection=Yes;'.format(dsn)

    def connect(self):
        if self.conn_type == 'alch':
            self.engine = mysql.create_engine(self.connstring)
        else:
            self.conn = pyodbc.connect(self.connstring)
            self.cursor = self.conn.cursor()
            self.conn.commit()

    def close(self):
        if self.conn_type == 'alch':
            self.engine.dispose()
        else:
            self.cursor.close()
            self.conn.close()

    def createsession(self):
        if self.conn_type == 'alch':
            self.engine = sessionmaker(bind=self.engine)
            self.engine = self.engine()
            self.engine._model_changes = {}
            self.session = True

    def createtable(self, dataframe, sqltable):
        if self.conn_type == 'alch' and not self.session:
            dataframe.to_sql(
                sqltable,
                self.engine,
                if_exists='replace',
            )

    def grabengine(self):
        if self.conn_type == 'alch':
            return self.engine
        else:
            return [self.cursor, self.conn]

    def upload(self, dataframe, sqltable):
        if self.conn_type == 'alch' and not self.session:
            mytbl = sqltable.split(".")

            if len(mytbl) > 1:
                dataframe.to_sql(
                    mytbl[1],
                    self.engine,
                    schema=mytbl[0],
                    if_exists='append',
                    index=True,
                    index_label='linenumber',
                    chunksize=1000
                )
            else:
                dataframe.to_sql(
                    mytbl[0],
                    self.engine,
                    if_exists='replace',
                    index=False,
                    chunksize=1000
                )

    def query(self, query):
        try:
            if self.conn_type == 'alch':
                obj = self.engine.execute(mysql.text(query))

                if obj._saved_cursor.arraysize > 0:
                    data = obj.fetchall()
                    columns = obj._metadata.keys

                    return pd.DataFrame(data, columns=columns)

            else:
                df = sql.read_sql(query, self.conn)
                return df

        except ValueError as a:
            print('\t[-] {} : SQL Query failed.'.format(a))
            pass

    def execute(self, query):
        try:
            if self.conn_type == 'alch':
                self.engine.execute(mysql.text(query))
            else:
                self.cursor.execute(query)

        except ValueError as a:
            print('\t[-] {} : SQL Execute failed.'.format(a))
            pass


def conn_chk():
    if 'Server' in settings.keys() and 'Database' in settings.keys():
        asql = SQLConnect('alch')
        asql.connect()
        engine = asql.grabengine()
        try:
            obj = engine.execute(mysql.text("SELECT 1 from sys.sysprocesses"))

            if obj._saved_cursor.arraysize > 0:
                engine.dispose()
                return True
            else:
                engine.dispose()
                del settings['Server']
                del settings['Database']
                return False

        except:
            engine.dispose()
            del settings['Server']
            del settings['Database']
            return False
    else:
        return False


def check_setting(setting_name):
    if setting_name not in settings.keys():
        print("Please type {} name:".format(setting_name))
        settings[setting_name] = input()


def init():
    global settings


def write_log(message, action='info'):
    filepath = os.path.join(settings['EventPath'],
                            "{} Event_Log.txt".format(datetime.datetime.now().__format__("%Y%m%d")))

    logging.basicConfig(filename=filepath,
                        level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

    print('{0} - {1} - {2}'.format(datetime.datetime.now(), action.upper(), message))

    if action == 'debug':
        logging.debug(message)
    elif action == 'info':
        logging.info(message)
    elif action == 'warning':
        logging.warning(message)
    elif action == 'error':
        logging.error(message)
    elif action == 'critical':
        logging.critical(message)


settings = dict()
if os.environ['PYTHONPATH']:
    for path in os.getenv('PYTHONPATH').split(os.pathsep):
        if not path == os.path.dirname(os.path.abspath(__file__)):
            settings['PythonPath'] = path
            break

if not settings['PythonPath']:
    settings['PythonPath'] = ''

settings['ErrPath'] = os.path.join(settings['PythonPath'], '02_Error')

if not os.path.exists(settings['ErrPath']):
    os.makedirs(settings['ErrPath'])

settings['EventPath'] = os.path.join(settings['PythonPath'], '01_Event_Logs')

if not os.path.exists(settings['EventPath']):
    os.makedirs(settings['EventPath'])

sfile = shelve.open(os.path.join(settings['PythonPath'], 'Settings'))

type(sfile)

for k, v in sfile.items():
    settings[k] = v

while not conn_chk():
    check_setting("Server")
    check_setting("Database")

for k, v in settings.items():
    sfile[k] = v

sfile.close()
