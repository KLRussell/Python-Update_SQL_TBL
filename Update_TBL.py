from Conn import SQLConnect
from Conn import write_log

import pandas as pd
import pathlib as pl
import os
import copy

ProcPath = os.path.join(os.path.abspath(__file__), '01_To_Process')


class ExcelToSQL:
    errors = []
    primary_key = None

    def __init__(self):
        self.asql = SQLConnect('alch')

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
                self.append_errors(table, data, 'Table {} in excel tab does not exist in the sql server'
                                   .format(table))
                return False
            else:
                return True
        else:
            self.append_errors(table, data, 'Table {} is not a proper (schema).(table) structure for excel tab'
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
                row = results.loc[results['Column_Name'] == col]

                if row.empty:
                    self.append_errors(table, data, 'Column {0} does not exist in {1}'
                                       .format(col, table))
                    return False
                elif row['Data_Type'] in ['xml', 'text', 'varchar', 'nvarchar', 'uniqueidentifier', 'nchar'
                    , 'geography', 'char', 'ntext'] and row['Character_Maximum_Length'] > 0:
                    myerr = data.loc[len(data[col]) > row['Character_Maximum_Length']]

                    self.append_errors(table, myerr,
                                       'Column {0} has {1} items that exceeds the limit percision for data type {2}'
                                       .format(col, len(myerr), row['Data_Type']))

                    if len(data) < 1:
                        return False
                elif row['Data_Type'] in ['varbinary', 'binary', 'bit', 'int', 'tinyint', 'smallint', 'bigint']:
                    myerr = data.loc[not data[col].str.isnumeric()]
                    self.append_errors(table, myerr,
                                       'Column {0} has {1} items that is not numeric for data type {2}'
                                       .format(col, len(myerr), row['Data_Type']))

                    if len(data) < 1:
                        return False

                    myerr = data.loc[data[col].str.isdigit()]
                    self.append_errors(table, myerr,
                                       'Column {0} has {1} items that has digits for data type {2}'
                                       .format(col, len(myerr), row['Data_Type']))

                    if len(data) < 1:
                        return False

                    if row['Data_Type'] in ['varbinary', 'binary']:
                        minnum = 1
                        maxnum = 8000
                    elif row['Data_Type'] == 'bit':
                        minnum = 0
                        maxnum = 1
                    elif row['Data_Type'] == 'tinyint':
                        minnum = 0
                        maxnum = 255
                    elif row['Data_Type'] == 'smallint':
                        minnum = -32768
                        maxnum = 32767
                    elif row['Data_Type'] == 'int':
                        minnum = -2147483648
                        maxnum = 2147483647
                    elif row['Data_Type'] == 'bigint':
                        minnum = -9223372036854775808
                        maxnum = 9223372036854775807
                    else:
                        minnum = 0
                        maxnum = 0

                    if row['Character_Maximum_Length'] > 0:
                        myerr = data.loc[len(data[col]) > row['Character_Maximum_Length']]
                    else:
                        myerr = data.loc[len(data[col]) < minnum]

                    self.append_errors(table, myerr,
                                       'Column {0} has {1} items that exceeds the minumum limit percision for data type {2}'
                                       .format(col, len(myerr), row['Data_Type']))

                    if len(data) < 1:
                        return False

                    if row['Character_Maximum_Length'] > 0:
                        myerr = data.loc[len(data[col]) > row['Character_Maximum_Length']]
                    else:
                        myerr = data.loc[len(data[col]) > maxnum]

                    self.append_errors(table, myerr,
                                       'Column {0} has {1} items that exceeds the maximum limit percision for data type {2}'
                                       .format(col, len(myerr), row['Data_Type']))

                    if len(data) < 1:
                        return False
                elif row['Data_Type'] in ['smalldatetime', 'date', 'datetime', 'datetime2', 'time']:
                    cleaned_df = pd.DataFrame()
                    cleaned_df['Date'] = pd.to_datetime(data[col], errors='coerce')
                    myerr = data.loc[cleaned_df.loc[cleaned_df['Date'].isnull()].index]

                    self.append_errors(table, myerr,
                                       'Column {0} has {1} items that isn''t in date/time format for data type {2}'
                                       .format(col, len(myerr), row['Data_Type']))

                    if len(data) < 1:
                        return False
                elif row['Data_Type'] in ['money', 'smallmoney', 'numeric', 'decimal', 'float', 'real']:
                    print('hi')
        else:
            self.append_errors(table, data, 'Unable to find table {} in INFORMATION_SCHEMA.COLUMNS table'
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
            for pk in results['Column_Name']:
                if pk in data.columns.tolist():
                    if self.primary_key:
                        self.primary_key = pk
                    else:
                        self.append_errors\
            (table, data, 'Columns {0} & {1} are Primary Keys. Please list only one Primary Key for tab {2}'
             .format(self.primary_key, pk, table))

                        return False
        else:
            self.append_errors(table, data, 'Table {} in SQL does not have a Primary Key. Unable to update records'
                               .format(table))

            return False

        if self.primary_key:
            return True
        else:
            self.append_errors(table, data, 'Tab {0} in excel has no Primary Key in tab. Please add one Primary Key as a column in this tab'.format(table))

            return False

    def process_file(self, myfile):
        xls_file = pd.ExcelFile(myfile)
        for table in xls_file.sheet_names:
            data = xls_file.parse(table)

            if self.validate_tab(table, data) and self.validate_data(table, data):
                print('success')

    def append_errors(self, table, df, errmsg):
        if not df.empty:
            write_log('{} Error(s) found. Appending to virtual list'.format(len(df.index)), 'warning')
            self.errors.append([table, copy.copy(df), errmsg])
            df.drop(df.index, inplace=True)


if __name__ == '__main__':
    dfs = []

    if not os.path.exists(ProcPath):
        os.makedirs(ProcPath)

    files = list(pl.Path(ProcPath).glob('*.xls*'))
    myobj = ExcelToSQL()

    for file in files:
        myobj.process_file(file)
