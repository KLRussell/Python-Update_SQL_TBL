from Global import grabobjs

import pandas as pd
import pathlib as pl
import os
import copy

globalobjs = grabobjs(os.path.dirname(os.path.abspath(__file__)))
ProcPath = os.path.join(os.path.dirname(os.path.abspath(__file__)), '01_To_Process')


class ExcelToSQL:
    errors = []
    primary_key = None

    def __init__(self):
        self.asql = SQLConnect('alch')
        self.asql.connect()

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
                self.append_errors(table, data, data, 'Table {} in excel tab does not exist in the sql server'
                                   .format(table))
                return False
            else:
                return True
        else:
            self.append_errors(table, data, data, 'Table {} is not a proper (schema).(table) structure for excel tab'
                               .format(table))
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
                    self.append_errors(table, data, data, 'Column {0} does not exist in {1}'
                                       .format(col, table))
                    return False
                elif row['Data_Type'][0] in \
                        ['xml', 'text', 'varchar', 'nvarchar', 'uniqueidentifier', 'nchar', 'geography', 'char', 'ntext']\
                        and str(row['Character_Maximum_Length'][0]).isnumeric():
                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if len(str(x)) > row['Character_Maximum_Length'][0] else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that exceeds the limit percision for data type {2}'
                                       .format(col, len(myerr), row['Data_Type'][0]))

                    if len(data) < 1:
                        return False
                elif row['Data_Type'][0] in ['varbinary', 'binary', 'bit', 'int', 'tinyint', 'smallint', 'bigint']:
                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(lambda x: True if str(x).isnumeric() else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that is not numeric for data type {2}'
                                       .format(col, len(myerr), row['Data_Type'][0]))

                    if len(data) < 1:
                        return False

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(lambda x: True if str(x).isdigit() else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that has digits for data type {2}'
                                       .format(col, len(myerr), row['Data_Type'][0]))

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

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that exceeds the minumum number size for data type {2}'
                                       .format(col, len(myerr), row['Data_Type'][0]))

                    if len(data) < 1:
                        return False

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if x > maxnum else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that exceeds the maximum number size for data type {2}'
                                       .format(col, len(myerr), row['Data_Type'][0]))

                    if len(data) < 1:
                        return False

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if len(str(x)) > row['Character_Maximum_Length'][0] else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that exceeds the precision size for data type {2}'
                                       .format(col, len(myerr), row['Data_Type'][0]))

                    if len(data) < 1:
                        return False
                elif row['Data_Type'][0] in ['smalldatetime', 'date', 'datetime', 'datetime2', 'time']:
                    cleaned_df = pd.DataFrame()
                    cleaned_df['Date'] = pd.to_datetime(data[col], errors='coerce')
                    myerr = data.loc[cleaned_df.loc[cleaned_df['Date'].isnull()].index].reset_index()

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that isn''t in date/time format for data type {2}'
                                       .format(col, len(myerr), row['Data_Type'][0]))

                    if len(data) < 1:
                        return False
                elif row['Data_Type'][0] in ['money', 'smallmoney', 'numeric', 'decimal', 'float', 'real']:
                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(lambda x: True if str(x).isnumeric() else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that is not numeric for data type {2}'
                                       .format(col, len(myerr), row['Data_Type'][0]))

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

                        self.append_errors(table, myerr, data,
                                           'Column {0} has {1} items that exceeds the minumum number size for data type {2}'
                                           .format(col, len(myerr), row['Data_Type'][0]))

                        if len(data) < 1:
                            return False

                        cleaned_df = pd.DataFrame()
                        cleaned_df[col] = data[col].map(
                            lambda x: True if x > maxnum else False)
                        myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                        self.append_errors(table, myerr, data,
                                           'Column {0} has {1} items that exceeds the maximum number size for data type {2}'
                                           .format(col, len(myerr), row['Data_Type'][0]))

                        if len(data) < 1:
                            return False

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if ('.' in str(x) and len(str(x).split('.')[0]) >
                                           row['Numeric_Precision'][0]) or ('.' not in str(x) and len(str(x)) >
                                                                            row['Numeric_Precision'][0]) else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that exceeds the numeric precision for data type {2}'
                                       .format(col, len(myerr), row['Data_Type'][0]))

                    if len(data) < 1:
                        return False

                    cleaned_df = pd.DataFrame()
                    cleaned_df[col] = data[col].map(
                        lambda x: True if ('.' in str(x) and len(str(x).split('.')[1]) >
                                           row['Numeric_Scale'][0]) or '.' not in str(x) else False)
                    myerr = data.loc[cleaned_df[cleaned_df[col].isnull()].index].reset_index()

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that exceeds the numeric scale for data type {2}'
                                       .format(col, len(myerr), row['Data_Type'][0]))

                    if len(data) < 1:
                        return False

                if row['Is_Nullable'][0] == 'NO':
                    myerr = data.loc[data[col].isnull()].reset_index()

                    self.append_errors(table, myerr, data,
                                       'Column {0} has {1} items that are null for data type {2} when null is not allowed'
                                       .format(col, len(myerr), row['Data_Type'][0]))

                    if len(data) < 1:
                        return False
        else:
            self.append_errors(table, data, data, 'Unable to find table {} in INFORMATION_SCHEMA.COLUMNS table'
                               .format(table))
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
            for pk in results['COLUMN_NAME'].tolist():
                if pk in data.columns.tolist():
                    if not self.primary_key:
                        self.primary_key = pk
                    else:
                        self.append_errors\
            (table, data, 'Columns {0} & {1} are Primary Keys. Please list only one Primary Key for tab {2}'
             .format(self.primary_key, pk, table))

                        return False
        else:
            self.append_errors(table, data, data, 'Table {} in SQL does not have a Primary Key. Unable to update records'
                               .format(table))

            return False

        if self.primary_key:
            return True
        else:
            self.append_errors(table, data, data, 'Tab {0} in excel has no Primary Key in tab. Please add one Primary Key as a column in this tab'.format(table))

            return False

    def update_tbl(self, table, data):
        self.asql.upload(data, 'UT_TMP')

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
            self.append_errors(table, myerr, data,
                               'Column {0} has {1} items that do not exist in table {2}'
                               .format(self.primary_key, len(myerr), table))

        if not data.empty:
            self.asql.execute('''
                update B
                set
                    {0}
                
                from UT_TMP As A
                left join {1} As B
                on
                    A.{2} = B.{2}
            '''.format(self.format_sql_set(data.columns.tolist()), table, self.primary_key))

        self.asql.execute('drop table UT_TMP')

    def format_sql_set(self, cols):
        myreturn = None
        for col in cols:
            if not col == self.primary_key:
                if myreturn:
                    myreturn = '{0}, B.{1} = A.{1}'.format(myreturn, col)
                else:
                    myreturn = 'B.{0} = A.{0}'.format(col)

        return myreturn

    def append_errors(self, table, err, df, errmsg):
        if not err.empty:
            write_log('{} Error(s) found. Appending to virtual list'.format(len(err.index)), 'warning')
            self.errors.append([table, copy.copy(err), errmsg])
            df.drop(err.index, inplace=True)

    def process_errs(self, data):
        print('processing errors')

    def close_sql(self):
        self.asql.close()


if __name__ == '__main__':
    dfs = []

    if not os.path.exists(ProcPath):
        os.makedirs(ProcPath)

    files = list(pl.Path(ProcPath).glob('*.xls*'))
    myobj = ExcelToSQL()

    for file in files:
        xls_file = pd.ExcelFile(file)

        for table in xls_file.sheet_names:
            data = xls_file.parse(table)

            if myobj.validate_tab(table, data) and myobj.validate_data(table, data):
                myobj.update_tbl(table, data)
                myobj.process_errs(data)

    myobj.close_sql()

