from Conn import SQLConnect
from Conn import write_log

import pandas as pd
import pathlib as pl
import os

ProcPath = os.path.join(os.path.abspath(__file__), '01_To_Process')


class ExcelToSQL:
    errors = []

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
        results = self.asql.query('''
            select
                Column_Name,
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
                elif row['Data_Type'] == 'int':

        else:
            self.append_errors(table, data, 'Unable to find table {} in INFORMATION_SCHEMA.COLUMNS table'
                               .format(table))
            return False

    def process_file(self, myfile):
        xls_file = pd.ExcelFile(myfile)
        for table in xls_file.sheet_names:
            data = xls_file.parse(table)

            if self.validate_tab(table, data) and self.validate_data(table, data):

    def append_errors(self, table, df, errmsg):
        if not df.empty:
            write_log('{} Error(s) found. Appending to virtual list'.format(len(df.index)), 'warning')
            self.errors.append([table, df, errmsg])


if __name__ == '__main__':
    dfs = []

    if not os.path.exists(ProcPath):
        os.makedirs(ProcPath)

    files = list(pl.Path(ProcPath).glob('*.xls*'))
    myobj = ExcelToSQL()

    for file in files:
        myobj.process_file(file)
