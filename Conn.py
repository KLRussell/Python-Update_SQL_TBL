from pandas.io import sql
from sqlalchemy.orm import sessionmaker
from urllib.parse import quote_plus

import pandas as pd
import sqlalchemy as mysql
import shelve
import pyodbc


class SQLConnect:
    session = False
    engine = None
    conn = None
    cursor = None

    def __init__(self, conn_type, dsn=None):
        self.conn_type = conn_type

        if conn_type == 'alch':
            self.connstring = self.alchconnstr(
                '{SQL Server Native Client 11.0}', 1433, sfile['Server'], sfile['Database'], 'mssql'
                )
        elif conn_type == 'sql':
            self.connstring = self.sqlconnstr(sfile['Server'], sfile['Database'])
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
    if 'Server' in sfile.keys() and 'Database' in sfile.keys():
        asql = SQLConnect('alch')
        engine = asql.grabengine()

        try:
            obj = engine.execute(mysql.text("SELECT VERSION()"))

            if obj._saved_cursor.arraysize > 0:
                engine.close()
                return True
            else:
                engine.close()
                del sfile['Server']
                del sfile['Database']
                return False

        except:
            engine.close()
            del sfile['Server']
            del sfile['Database']
            return False
    else:
        return False


def check_setting(setting_name):
    if setting_name not in sfile.keys():
        print("Please type {} name:".format(setting_name))
        sfile[setting_name] = input()


sfile = shelve.open('Settings')

type(sfile)

while not conn_chk:
    check_setting("Server")
    check_setting("Database")

sfile.close()
