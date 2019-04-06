from Global import grabobjs
from Global import ShelfHandle
from time import sleep

import pandas as pd
import pathlib as pl
import os
import copy
import win32api
import win32con
import win32security

CurrDir = os.path.dirname(os.path.abspath(__file__))
ProcDir = os.path.join(CurrDir, '02_To_Process')
ErrDir = os.path.join(CurrDir, '03_Errors')
PreserveDir = os.path.join(CurrDir, '04_Preserve')
Global_Objs = grabobjs(CurrDir)
Preserve_Obj = None


class ExcelToSQL:
    primary_key = None

    def __init__(self, mode):
        self.mode = mode
        self.errors_obj = Global_Objs['Errors']
        self.asql = Global_Objs['SQL'].connect('alch')

    def validate_tab(self, table, data):
        splittable = table.split('.')

        if len(splittable) == 2:
            results = self.asql.query('''
                select 1
                from information_schema.tables
                where
                    table_schema = '{0}'
                        and
                    table_name = '{1}'
            '''.format(splittable[0], splittable[1]))

            if results.empty:
                mylist = [copy.copy(table), copy.copy(data),
                          'Table {} in excel tab does not exist in the sql server'.format(table)]
                self.errors_obj.append_errors(mylist)

                return False
            else:
                return True
        else:
            mylist = [copy.copy(table), copy.copy(data),
                      'Table {} is not a proper (schema).(table) structure for excel tab'.format(table)]
            self.errors_obj.append_errors(mylist)

            return False

    def validate_data(self, table, data):
        self.primary_key = None

        results = self.asql.query('''
            select
                Column_Name,
                Is_Nullable,
                Data_Type,
                Character_Maximum_Length,
                Numeric_Precision,
                Numeric_Scale

            from INFORMATION_SCHEMA.COLUMNS

            where
                TABLE_SCHEMA = '{0}'
                    and
                TABLE_NAME = '{1}'
        '''.format(table.split('.')[0], table.split('.')[1]))

        if not results.empty:
            for col in data.columns.tolist():
                row = results.loc[results['Column_Name'] == col].reset_index()

                if row.empty:
                    mylist = [copy.copy(table), copy.copy(data), 'Column {0} does not exist in {1}'.format(col, table)]
                    self.errors_obj.append_errors(mylist)

                    return False
                elif row['Data_Type'][0] in \
                        ['xml', 'text', 'varchar', 'nvarchar', 'uniqueidentifier', 'nchar', 'geography', 'char', 'ntext']\
                        and str(row['Character_Maximum_Length'][0]).isnumeric():
                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if len(str(x)) > row['Character_Maximum_Length'][0] else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that exceeds the limit percision for data type {2}'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False
                elif row['Data_Type'][0] in ['varbinary', 'binary', 'bit', 'int', 'tinyint', 'smallint', 'bigint']:
                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(lambda x: True if str(x).isnumeric() else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that is not numeric for data type {2}'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(lambda x: True if str(x).isdigit() else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that has digits for data type {2}'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False

                    if row['Data_Type'][0] in ['varbinary', 'binary']:
                        minnum = 1
                        maxnum = 8000
                    elif row['Data_Type'][0] == 'bit':
                        minnum = 0
                        maxnum = 1
                    elif row['Data_Type'][0] == 'tinyint':
                        minnum = 0
                        maxnum = 255
                    elif row['Data_Type'][0] == 'smallint':
                        minnum = -32768
                        maxnum = 32767
                    elif row['Data_Type'][0] == 'int':
                        minnum = -2147483648
                        maxnum = 2147483647
                    elif row['Data_Type'][0] == 'bigint':
                        minnum = -9223372036854775808
                        maxnum = 9223372036854775807
                    else:
                        minnum = 0
                        maxnum = 0

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if x < minnum else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that exceeds the minumum number size for data type {2}'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if x > maxnum else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that exceeds the maximum number size for data type {2}'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if len(str(x)) > row['Character_Maximum_Length'][0] else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that exceeds the precision size for data type {2}'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False
                elif row['Data_Type'][0] in ['smalldatetime', 'date', 'datetime', 'datetime2', 'time']:
                    cleaned_df = pd.DataFrame()
                    cleaned_df['Date'] = pd.to_datetime(data[col], errors='coerce')
                    myerr = data.loc[cleaned_df.loc[cleaned_df['Date'].isnull()].index].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that isn''t in date/time format for data type {2}'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False
                elif row['Data_Type'][0] in ['money', 'smallmoney', 'numeric', 'decimal', 'float', 'real']:
                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(lambda x: True if str(x).isnumeric() else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that is not numeric for data type {2}'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False

                    if row['Data_Type'][0] == 'money':
                        minnum = -922337203685477.5808
                        maxnum = 922337203685477.5807
                    elif row['Data_Type'][0] == 'smallmoney':
                        minnum = -214748.3648
                        maxnum = 214748.3647
                    elif row['Data_Type'][0] in ['decimal', 'numeric']:
                        minnum = -10 ** 38 + 1
                        maxnum = 10 ** 38 - 1

                    if row['Data_Type'][0] in ['money', 'smallmoney', 'decimal', 'numeric']:
                        cleaned_df = pd.DataFrame()
                        cleaned_df[col] = data[col].map(
                            lambda x: True if x < minnum else False)
                        myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                        if not myerr.empty:
                            mylist = [copy.copy(table), copy.copy(myerr),
                                      'Column {0} has {1} items that exceeds the minumum number size for data type {2}'
                                          .format(col, len(myerr), row['Data_Type'][0])]
                            self.errors_obj.append_errors(mylist)
                            self.errors_obj.trim_df(data, myerr)

                        if len(data) < 1:
                            return False

                        cleaned_df = pd.DataFrame()
                        cleaned_df[col] = data[col].map(
                            lambda x: True if x > maxnum else False)
                        myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                        if not myerr.empty:
                            mylist = [copy.copy(table), copy.copy(myerr),
                                      'Column {0} has {1} items that exceeds the maximum number size for data type {2}'
                                          .format(col, len(myerr), row['Data_Type'][0])]
                            self.errors_obj.append_errors(mylist)
                            self.errors_obj.trim_df(data, myerr)

                        if len(data) < 1:
                            return False

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if ('.' in str(x) and len(str(x).split('.')[0]) >
                                           row['Numeric_Precision'][0]) or ('.' not in str(x) and len(str(x)) >
                                                                            row['Numeric_Precision'][0]) else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that exceeds the numeric precision for data type {2}'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if ('.' in str(x) and len(str(x).split('.')[1]) >
                                           row['Numeric_Scale'][0]) or '.' not in str(x) else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that exceeds the numeric scale for data type {2}'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False

                if row['Is_Nullable'][0] == 'NO':
                    myerr = data.loc[data[col].isnull()].reset_index()

                    if not myerr.empty:
                        mylist = [copy.copy(table), copy.copy(myerr),
                                  'Column {0} has {1} items that are null for data type {2} when null is not allowed'
                                      .format(col, len(myerr), row['Data_Type'][0])]
                        self.errors_obj.append_errors(mylist)
                        self.errors_obj.trim_df(data, myerr)

                    if len(data) < 1:
                        return False
        else:
            mylist = [copy.copy(table), copy.copy(data),
                      'Unable to find table {} in INFORMATION_SCHEMA.COLUMNS table'.format(table)]
            self.errors_obj.append_errors(mylist)

            return False

        results = self.asql.query('''
            SELECT
                K.COLUMN_NAME
            
            FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS AS C
            INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS K
            ON
                C.TABLE_NAME = K.TABLE_NAME
                    AND
                C.CONSTRAINT_CATALOG = K.CONSTRAINT_CATALOG
                    AND
                C.CONSTRAINT_SCHEMA = K.CONSTRAINT_SCHEMA
                    AND
                C.CONSTRAINT_NAME = K.CONSTRAINT_NAME
            
            WHERE
                K.CONSTRAINT_SCHEMA = '{0}'
                    AND
                K.TABLE_NAME = '{1}'
                    AND
                C.CONSTRAINT_TYPE = 'PRIMARY KEY'
        '''.format(table.split('.')[0], table.split('.')[1]))

        if not results.empty:
            if self.mode:
                mylist = [copy.copy(table), copy.copy(data),
                          'Tab {} in excel spreadsheet has a Primary Key when trying to insert records'
                              .format(table)]
                self.errors_obj.append_errors(mylist)

                return False
            else:
                for pk in results['COLUMN_NAME'].tolist():
                    if pk in data.columns.tolist():
                        if not self.primary_key:
                            self.primary_key = pk
                        else:
                            mylist = [copy.copy(table), copy.copy(data),
                                      'Columns {0} & {1} are Primary Keys. Please list only one Primary Key for tab {2}'
                                          .format(self.primary_key, pk, table)]
                            self.errors_obj.append_errors(mylist)

                            return False
        elif self.mode:
            return True
        else:
            mylist = [copy.copy(table), copy.copy(data),
                      'Table {} in SQL does not have a Primary Key. Unable to update records'
                          .format(table)]
            self.errors_obj.append_errors(mylist)

            return False

        if self.primary_key:
            return True
        else:
            mylist = [copy.copy(table), copy.copy(data),
                      'Tab {0} in excel has no Primary Key in tab. Please add one Primary Key as a column in this tab'
                          .format(table)]
            self.errors_obj.append_errors(mylist)

            return False

    def update_tbl(self, table, data):
        self.asql.upload(data, 'UT_TMP')

        if self.mode:
            self.asql.execute('''
                insert into {0}
                (
                    {1}
                )
                select
                    {1}
                
                from UT_TMP
            '''.format(table, self.format_sql_set(data.columns.tolist())))
        else:
            results = self.asql.query('''
                select
                    A.{0}
                from UT_TMP As A
                left join {1} As B
                on
                    A.{0} = B.{0}
                
                where
                    B.{0} is null
            '''.format(self.primary_key, table))

            if not results.empty:
                myerr = data[data[self.primary_key].isin(results[self.primary_key])]

                if not myerr.empty:
                    mylist = [copy.copy(table), copy.copy(myerr),
                              'Column {0} has {1} items that do not exist in table {2}'
                                  .format(self.primary_key, len(myerr), table)]
                    self.errors_obj.append_errors(mylist)
                    self.errors_obj.trim_df(data, myerr)

            if not data.empty:
                mydf = self.asql.query('''
                    select
                        B.{2},
                        {0}
                        
                    from UT_TMP As A
                    inner join {1} As B
                    on
                        A.{2} = B.{2}
                '''.format(self.format_sql_set(data.columns.tolist(), 'B.'), table, self.primary_key))
                self.shelf_old(mydf)

                self.asql.execute('''
                    update B
                    set
                        {0}
                    
                    from UT_TMP As A
                    inner join {1} As B
                    on
                        A.{2} = B.{2}
                '''.format(self.format_sql_set(data.columns.tolist()), table, self.primary_key))

        self.asql.execute('drop table UT_TMP')

    def format_sql_set(self, cols, prefix=None):
        myreturn = None
        for col in cols:
            if not col == self.primary_key:
                if self.mode:
                    if myreturn:
                        myreturn = '{0}, {1}'.format(myreturn, col)
                    else:
                        myreturn = '{0}'.format(col)
                elif prefix:
                    if myreturn:
                        myreturn = '{0}, {1}{2}'.format(myreturn, prefix, col)
                    else:
                        myreturn = '{0}{1}'.format(prefix, col)
                else:
                    if myreturn:
                        myreturn = '{0}, B.{1} = A.{1}'.format(myreturn, col)
                    else:
                        myreturn = 'B.{0} = A.{0}'.format(col)

        return myreturn

    @staticmethod
    def shelf_old(df):
        today = datetime.datetime.now().__format__("%Y%m%d")
        mylist = Preserve_Obj.grab_item(today)

        if mylist:
            Preserve_Obj.del_item(today)
            mylist.append(df)
            Preserve_Obj.add_item(today, mylist)
        else:
            mylist = [].append(df)
            Preserve_Obj.add_item(today, mylist)

    def process_errs(self, file):
        myerrs = self.errors_obj.grab_errors()

        if myerrs:
            errmsgs = []
            sd = win32security.GetFileSecurity(file, win32security.OWNER_SECURITY_INFORMATION)
            owner_sid = sd.GetSecurityDescriptorOwner()
            creator, domain, type = win32security.LookupAccountSid(None, owner_sid)

            Global_Objs['Event_Log'].write_log('Appending errors into {0} ({1})'.format(file, creator), 'error')

            with pd.ExcelWriter(file) as writer:
                for err in myerrs:
                    errmsgs.append(('%s\\%s' % (domain, creator), err[1], err[2]))
                    err[1].to_excel(writer, sheet_name=err[0])

                df = pd.DataFrame(errmsgs, ['File_Creator_Name', 'Tab_Name', 'Errors'])
                df.to_excel(writer, sheet_name='Error_Details')

    def close_sql(self):
        self.asql.close()


def check_for_updates():
    f = list(pl.Path(ProcDir).glob('Update_*.xls*'))

    if f:
        return [f, False]

    f = list(pl.Path(ProcDir).glob('Insert_*.xls*'))

    if f:
        return [f, True]


def process_updates(info):
    myobj = ExcelToSQL(info[1])

    for file in info[0]:
        Global_Objs['Event_Log'].write_log('Processing file {}'.format(file))
        xls_file = pd.ExcelFile(file)

        for tbl in xls_file.sheet_names:
            df = xls_file.parse(tbl)
            Global_Objs['Event_Log'].write_log('Validating tab {} for errors'.format(tbl))

            if myobj.validate_tab(tbl, df) and myobj.validate_data(tbl, df):
                if info[1]:
                    Global_Objs['Event_Log'].write_log('Inserting {0} items into {1}'.format(len(df), tbl))
                else:
                    Global_Objs['Event_Log'].write_log('Updating {0} items in {1}'.format(len(df), tbl))
                myobj.update_tbl(tbl, df)
                myobj.process_errs(file)

    myobj.close_sql()
    del myobj


if __name__ == '__main__':
    if not os.path.exists(ProcDir):
        os.makedirs(ProcDir)

    if not os.path.exists(ErrDir):
        os.makedirs(ErrDir)

    if not os.path.exists(PreserveDir):
        os.makedirs(PreserveDir)

    Preserve_Obj = ShelfHandle(os.path.join(PreserveDir, 'Data_Locker'))
    has_updates = check_for_updates()

    if has_updates:
        Global_Objs['Event_Log'].write_log('Found {} files to process'.format(len(has_updates[0])))
        process_updates(has_updates)
    else:
        Global_Objs['Event_Log'].write_log('Found no files to process', 'warning')

    os.system('pause')
