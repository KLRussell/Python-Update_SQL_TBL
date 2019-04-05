from pandas.io import sql
from sqlalchemy.orm import sessionmaker
from urllib.parse import quote_plus

import pandas as pd
import sqlalchemy as mysql
import shelve
import pyodbc
import os
import datetime
import logging


def grabobjs(scriptdir):
    myobjs = dict()
    myobjs['Settings'] = SettingsHandle(scriptdir)
    myobjs['Event_Log'] = LogHandle(myobjs['Settings'])
    myobjs['SQL'] = SQLHandle(myobjs['Settings'])
    myobjs['Errors'] = ErrHandle(myobjs['Event_Log'])

    return myobjs


class SettingsHandle:
    settings = dict()
    settingspath = None

    def __init__(self, scriptdir):
        self.load_script_settings(scriptdir)
        sfile = shelve.open(self.settingspath)
        type(sfile)

        for k, v in sfile.items():
            self.settings[k] = v

        sfile.close()

    def load_script_settings(self, scriptdir):
        if scriptdir and os.path.exists(scriptdir):
            scriptfilepath = os.path.join(scriptdir, 'Script_Settings')

            sfile = shelve.open(scriptfilepath)
            type(sfile)

            while 'General_Settings_Path' not in sfile.keys():
                print("Please input general settings path:")

                sfile['General_Settings_Path'] = input()

                if not sfile['General_Settings_Path'] or not os.path.exists(sfile['General_Settings_Path']):
                    del sfile['General_Settings_Path']

            self.settingspath = sfile['General_Settings_Path']
            sfile.close()
        else:
            raise Exception('Invalid scriptpath path')

    def find_setting(self, setting_name):
        if setting_name and setting_name in self.settings.keys():
            return self.settings[setting_name]

    def remove_setting(self, setting_name):
        if setting_name and setting_name in self.settings.keys():
            del self.settings[setting_name]
            sfile = shelve.open(self.settingspath)
            type(sfile)
            del sfile[setting_name]
            sfile.close()

    def append_settings(self, newsettings):
        if type(newsettings) == 'dict' and len(newsettings) > 0:
            sfile = shelve.open(self.settingspath)
            type(sfile)

            for k, v in newsettings.items():
                if k not in self.settings.keys():
                    self.settings[k] = v

            for k, v in newsettings.items():
                if k not in sfile.keys():
                    sfile[k] = v

            sfile.close()

    def new_setting(self, setting_name):
        if setting_name and setting_name not in self.settings.keys():
            print("Please type {} value:".format(setting_name))
            self.settings[setting_name] = input()
            if not self.settings[setting_name]:
                self.new_setting(setting_name)

            sfile = shelve.open(self.settingspath)
            type(sfile)
            sfile[setting_name] = self.settings[setting_name]
            sfile.close()


class LogHandle:
    def __init__(self, settingsobj):
        if settingsobj:
            self.settingsobj = settingsobj

            while not self.settingsobj.find_setting('Event_Log_Path'):
                self.settingsobj.new_setting('Event_Log_Path')

                if not os.path.exists(self.settingsobj.find_setting('Event_Log_Path')):
                    self.settingsobj.remove_setting('Event_Log_Path')
        else:
            raise Exception('Settings object was not passed through')

    def write_log(self, message, action='info'):
        filepath = os.path.join(self.settingsobj.find_setting('Event_Log_Path'),
                                "{0}_{1}_Log.txt".format(datetime.datetime.now().__format__("%Y%m%d"), os.path
                                                         .basename(os.path.dirname(os.path.abspath(__file__)))))

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


class SQLHandle:
    conn_type = None
    conn_str = None
    session = False
    engine = None
    conn = None
    cursor = None

    def __init__(self, settingsobj):
        if settingsobj:
            self.settingsobj = settingsobj
        else:
            raise Exception('Settings object not included in parameter')

    def create_conn_str(self, server=None, database=None, dsn=None):
        if self.conn_type == 'alch':
            p = quote_plus(
                'DRIVER={};PORT={};SERVER={};DATABASE={};Trusted_Connection=yes;'
                    .format('{SQL Server Native Client 11.0}', '1433', server, database))

            self.conn_str = '{}+pyodbc:///?odbc_connect={}'.format('mssql', p)
        elif self.conn_type == 'sql':
            self.conn_str = 'driver={0};server={1};database={2};autocommit=True;Trusted_Connection=yes'\
                .format('{SQL Server}', server, database)
        elif self.conn_type == 'dsn':
            self.conn_str = 'DSN={};DATABASE=default;Trusted_Connection=Yes;'.format(dsn)
        else:
            raise Exception('Invalid conn_type specified')

    def val_settings(self):
        if self.conn_type in ['alch', 'sql']:
            if not self.settingsobj.find_setting('Server') and not self.settingsobj.find_setting('Database'):
                self.settingsobj.new_setting('Server')
                self.settingsobj.new_setting('Database')

            self.create_conn_str(server=self.settingsobj.find_setting('Server')
                                 , database=self.settingsobj.find_setting('Database'))

        elif self.conn_type == 'dsn':
            if not self.settingsobj.find_setting('DSN'):
                self.settingsobj.new_setting('DSN')

            self.create_conn_str(dsn=self.settingsobj.find_setting('DSN'))

    def conn_chk(self):
        exit_loop = False

        while exit_loop:
            self.val_settings()
            myquery = "SELECT 1 from sys.sysprocesses"
            self.connect()

            try:
                if self.conn_type == 'alch':
                    obj = self.engine.execute(mysql.text(myquery))

                    if obj._saved_cursor.arraysize > 0:
                        exit_loop = True
                    else:
                        self.settingsobj.remove_setting('Server')
                        self.settingsobj.remove_setting('Database')
                        print('Error! Server & Database combination are incorrect!')

                else:
                    df = sql.read_sql(myquery, self.conn)
                    if len(df) > 0:
                        exit_loop = True
                    else:
                        if self.conn_type == 'sql':
                            self.settingsobj.remove_setting('Server')
                            self.settingsobj.remove_setting('Database')
                            print('Error! Server & Database combination are incorrect!')
                        else:
                            self.settingsobj.remove_setting('DSN')
                            print('Error! DSN is incorrect!')

                self.close()

            except ValueError as a:
                if self.conn_type in ['alch', 'sql']:
                    self.settingsobj.remove_setting('Server')
                    self.settingsobj.remove_setting('Database')
                    print('Error! Server & Database combination are incorrect!')
                else:
                    self.settingsobj.remove_setting('DSN')
                    print('Error! DSN is incorrect!')

                self.close()

    def connect(self, conn_type):
        self.conn_type = conn_type
        self.conn_chk()

        if self.conn_type == 'alch':
            self.engine = mysql.create_engine(self.conn_str)
        else:
            self.conn = pyodbc.connect(self.conn_str)
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


class ErrHandle:
    def __init__(self, logobj):
        self.errors = dict()
        self.logobj = logobj

    @staticmethod
    def trim_df(df_to_trim, df_to_compare):
        if type(df_to_trim) == 'DataFrame' and type(df_to_compare) == 'DataFrame' and not df_to_trim.empty\
                and not df_to_compare.empty:
            df_to_trim.drop(df_to_compare.index, inplace=True)

    @staticmethod
    def concat_dfs(df_list):
        if type(df_list) == 'list' and len(df_list) > 0:
            dfs = []

            for df in df_list:
                if type(df) == 'DataFrame':
                    dfs.append(df)

            if len(dfs) > 0:
                return pd.concat(dfs, ignore_index=True, sort=False).drop_duplicates().reset_index(drop=True)

    def append_errors(self, err_items, key=None):
        if type(err_items) == 'list' and not err_items:
            self.logobj.write_log('Error(s) found. Appending to virtual list', 'warning')

            if key and key in self.errors.keys:
                self.errors[key] = self.errors[key].append(err_items)
            elif key:
                self.errors[key] = [].append(err_items)
            elif 'default' in self.errors.keys:
                self.errors['default'] = self.errors['default'].append(err_items)
            else:
                self.errors['default'] = [].append(err_items)

    def grab_errors(self, key=None):
        if key and key in self.errors.keys:
            mylist = self.errors[key]
            del self.errors[key]
            return mylist
        elif not key and 'default' in self.errors.keys:
            mylist = self.errors['default']
            del self.errors['default']
            return mylist
        else:
            return None
