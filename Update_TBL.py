import pandas as pd
import pathlib as pl
import os

CurrPath = os.path.dirname(os.path.realpath(__file__))
ProcPath = os.path.join(CurrPath, '01_To_Process')
ErrPath = os.path.join(CurrPath, '02_Error')


def process_file(myfile):
    xls_file = pd.ExcelFile(myfile)
    for table in xls_file.sheet_names:
        data = xls_file.parse(table)

        print(data)


if __name__ == '__main__':
    dfs = []

    files = list(pl.Path(ProcPath).glob('*.xls*'))

    for file in files:
        process_file(file)
