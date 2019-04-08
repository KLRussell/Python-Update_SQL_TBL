#############################################################################################################
					***READ ME*** - Update_TBL
					      by Kevin Russell
#############################################################################################################

- The purpose of this project is to make the process of small-scale SQL table updates and small-scale SQL
table inserts more efficient, organized, systematic, and easier.

- This script will check validity of SQL table, user inputted table columns, table data types, table data
percision/size, table column non-null setting, and existence of table primary key ids. Additionally, the
script will check whether user provided multiple primary key columns.

- Items that do not pass the above checks will be processed into error excel spreadsheets that is exported
into the error log directory '03_Errors' located in the script's filepath. Tab Error_Details, in the error
file, will list the problematic issues with the other tabs.

- Updated items appended to a SQL table will have the old values and table's primary key appended to a shelf
file. This shelf file is stored under '04_Preserve' under current filepath. Data is appended to a dict
datatype that is under the key of the current datetime of YYYYMMDD.

- Script will log all actions it performs in print and in file, which is located in '01_Event_Logs' within
script's filepath.

- Script will check table if table has an edit date column, and script will automatically update that column
with the current datetime.

- A secondary script will export items from locker into an excel spreadsheet


INSTALLATION:
	- Global.py can be placed in one of the PYTHONPATH directories

	- Update_TBL.py needs to be placed in a new folder and you can place that folder where you see fit

SETUP:
	- Run Update_TBL.py

	- A prompt will ask for you to provide a directory to store the general settings. Generally you can
	use the script's pathway, but in other cases you may want to choose a different location. I like to
	keep my settings in a PYTHONPATH directory.

	- A prompt will inquire you to input your Server address/name and Database name. This will be saved
	in your general settings file.

HOW IT WORKS:
	- Format excel spreadsheet into the following format:
		* Filename should start with either 'Insert_' or 'Update_'

		* Tab must be named as [schema].[table name] (File can have multiple tabs)

		* First row must have one primary key and the column names of what you plan to update

		* Remaining cells below the column names is either the id of the primary key or new data
		that replaces the old data in SQL

	- Place excel spreadsheet into '02_Process'

	- Execute Update_TBL.py

GRAB FROM PRESERVE LOCKER:
	- Run Update_TBL_Locker.py

	- Select a date from Preserve Locker to export into excel spreadsheet

	- Data will be exported into the '04_Preserve\Data_Locker_Export' directory of your script directory

	- Tab 'Append_Details' show details of each tab for that day in the Locker

