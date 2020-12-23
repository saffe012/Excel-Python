'''
Module of 'excel.py' that contains global functions.
Matt Saffert
1-9-2020
'''

import pyodbc
import re
from excel_constants import *
import subprocess
import sys
import pandas as pd
import global_gui as gui


def validateData(worksheet):
    '''Validates the data in the passed in worksheet. The data comes from the 6th
    row and on in an Excel spreadsheet

    :param1 worksheet: pandas.core.frame.DataFrame

    :return: bool
    '''

    valid_template = True

    for row in range(START_OF_DATA_ROWS_INDEX, len(worksheet) - 1):
        for i in range(len(worksheet.loc['info'])):
            if pd.isnull(worksheet.iloc[row][i]) and (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
                valid_template = False
                gui.createPopUpBox(
                    'You have not entered a value in cell ' + getExcelCellToInsertInto(i, row) + ' where one is required')
    blank_last_row = True
    for i in range(len(worksheet.iloc[len(worksheet) - 1])):
        if not (pd.isnull(worksheet.iloc[len(worksheet) - 1][i])):
            blank_last_row = False
    if not blank_last_row:
        for i in range(len(worksheet.iloc[len(worksheet) - 1])):
            if pd.isnull(worksheet.iloc[len(worksheet) - 1][i]) and (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
                valid_template = False
                gui.createPopUpBox(
                    'You have not entered a value in cell ' + getExcelCellToInsertInto(i, len(worksheet) - 1) + ' where one is required')

    return valid_template


def validateWorksheetSQL(worksheet):
    '''Validates the data in the passed in worksheet based on a SQL table from an
    open SQL connection.

    :param1 worksheet: pandas.core.frame.DataFrame

    :return: bool
    '''

    valid_template = True

    tables, cursor, sql_database_name = connectToSQLServer()
    if worksheet.loc['info'][0] == None or worksheet.loc['info'][0] not in tables:
        valid_template = False
        gui.createPopUpBox(
            'You have not specified a valid SQL table name in cell "A1"')
        gui.createPopUpBox(
            'Cannot continue SQL validation.')
        return valid_template

    if worksheet.loc['info'][1] not in TYPE_OF_SCRIPTS_AVAILABLE:
        valid_template = False
        gui.createPopUpBox(
            'You have not specified a valid script type in cell "B1"')

    sql_column_names, sql_column_types, column_is_nullable, column_is_identity = getSQLTableInfo(
        worksheet.loc['info'][0], cursor)

    for i in range(len(worksheet.loc['names'])):
        if (worksheet.loc['names'][i] == None or worksheet.loc['names'][i] not in sql_column_names) and (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
            valid_template = False
            gui.createPopUpBox(
                'You have not entered a column name where one is required in cell ' + getExcelCellToInsertInto(i, COLUMN_NAMES_ROW_INDEX))

    for i in range(len(worksheet.loc['types'])):
        type = re.sub("[\(\[].*?[\)\]]", "", str(worksheet.loc['types'][i]))
        if type not in SQL_STRING_TYPE and type not in SQL_NUMERIC_TYPE and type not in SQL_DATETIME_TYPE and type not in SQL_OTHER_TYPE:
            if (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
                valid_template = False
                gui.createPopUpBox(
                    'You have not entered a supported SQL type where one is required in cell ' + getExcelCellToInsertInto(i, COLUMN_DATA_TYPE_ROW_INDEX))
        column_name = worksheet.loc['names'][i]
        if column_name in sql_column_names:
            sql_name_index = sql_column_names.index(column_name)
            if type != sql_column_types[sql_name_index]:
                valid_template = False
                gui.createPopUpBox(
                    'The type in your spreadsheet for ' + column_name + ', does not match the type of the column in SQL in cell ' + getExcelCellToInsertInto(i, COLUMN_DATA_TYPE_ROW_INDEX))

    for i in range(len(worksheet.loc['include'])):
        if worksheet.loc['include'][i] != None and worksheet.loc['include'][i] != 'include':
            valid_template = False
            gui.createPopUpBox(
                'You have not entered an invalid string in cell ' + getExcelCellToInsertInto(i, INCLUDE_ROW_INDEX) + '. Valid string for row 4 is "include" or leave blank')
        if worksheet.loc['info'][1] != 'delete':
            if column_is_identity[i] == 0:
                # if script type is insert, and column cannot be null then automatically select
                if column_is_nullable[i] == 'NO' and worksheet.loc['info'][1] not in ('select', 'update'):
                    if worksheet.loc['include'][i] != 'include':
                        valid_template = False
                        gui.createPopUpBox(
                            'You have entered an invalid string in cell ' + getExcelCellToInsertInto(i, INCLUDE_ROW_INDEX) + '. This column must be included')
            else:  # column is identity column so cannot be updated or inserted into.
                # insert/update on identity column is NOT allowed
                if worksheet.loc['info'][1] != 'select':
                    if worksheet.loc['include'][i] == 'include':
                        valid_template = False
                        gui.createPopUpBox(
                            'You have entered an invalid string in cell ' + getExcelCellToInsertInto(i, INCLUDE_ROW_INDEX) + '. This column cannot be included')

    for i in range(len(worksheet.loc['where'])):
        if worksheet.loc['where'][i] != None and worksheet.loc['where'][i] != 'where':
            valid_template = False
            gui.createPopUpBox(
                'You have not entered an invalid string in a cell in cell ' + getExcelCellToInsertInto(i, WHERE_ROW_INDEX) + '. Valid string for row 5 is "where" or leave blank')

    return validateData(worksheet) and valid_template


def validateWorksheetGeneric(worksheet):
    '''Validates the data in the passed in worksheet based on a generic SQL table.

    :param1 worksheet: pandas.core.frame.DataFrame

    :return: bool
    '''

    valid_template = True

    if pd.isnull(worksheet.loc['info'][0]):
        valid_template = False
        gui.createPopUpBox(
            'You have not specified a SQL table name in cell "A1"')
    if worksheet.loc['info'][1] not in TYPE_OF_SCRIPTS_AVAILABLE:
        valid_template = False
        gui.createPopUpBox(
            'You have not specified a valid script type in cell "B1"')

    for i in range(len(worksheet.loc['names'])):
        if pd.isnull(worksheet.loc['names'][i]) and (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
            valid_template = False
            gui.createPopUpBox(
                'You have not entered a column name where one is required in cell ' + getExcelCellToInsertInto(i, COLUMN_NAMES_ROW_INDEX))

    for i in range(len(worksheet.loc['types'])):
        type = re.sub("[\(\[].*?[\)\]]", "",
                      str(worksheet.loc['types'][i]))
        if type not in SQL_STRING_TYPE and type not in SQL_NUMERIC_TYPE and type not in SQL_DATETIME_TYPE and type not in SQL_OTHER_TYPE:
            if (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
                valid_template = False
                gui.createPopUpBox(
                    'You have not entered a supported SQL type where one is required in cell ' + getExcelCellToInsertInto(i, COLUMN_DATA_TYPE_ROW_INDEX))

    for i in range(len(worksheet.loc['include'])):
        if not (pd.isnull(worksheet.loc['include'][i])) and worksheet.loc['include'][i] != 'include':
            valid_template = False
            gui.createPopUpBox(
                'You have not entered an invalid string in cell ' + getExcelCellToInsertInto(i, INCLUDE_ROW_INDEX) + '. Valid string for row 4 is "include" or leave blank')

    for i in range(len(worksheet.loc['where'])):
        if not (pd.isnull(worksheet.loc['where'][i])) and worksheet.loc['where'][i] != 'where':
            valid_template = False
            gui.createPopUpBox(
                'You have not entered an invalid string in a cell in cell ' + getExcelCellToInsertInto(i, WHERE_ROW_INDEX) + '. Valid string for row 5 is "where" or leave blank')

    return validateData(worksheet) and valid_template


def validWorksheet(worksheet, validate_with_sql, title):
    '''Calls the correct function to validate the passed worksheet based on
    whether a user wants to connect to SQL or not.

    :param1 worksheet: pandas.core.frame.DataFrame
    :param2 validate_with_sql: str
    :param3 title: str

    :return: bool
    '''

    description = "Would you like to validate/create scripts for " + \
        title + " worksheet?"
    yes = "Yes"
    no = "No"
    write_script_for, additional_box_val = gui.createYesNoBox(
        description, yes, no, additional_box=True)

    valid_template = True
    if validate_with_sql == 'Generic':
        if write_script_for == yes:  # if the user says to write scripts for this sheet
            valid_template = validateWorksheetGeneric(
                worksheet) and valid_template
        else:
            valid_template = False
            gui.createPopUpBox(
                'Validation failed. Scripts will not be written for ' + title)

    elif validate_with_sql == 'SQL':
        if write_script_for == yes:  # if the user says to write scripts for this sheet
            valid_template = validateWorksheetSQL(worksheet) and valid_template
        else:
            valid_template = False
            gui.createPopUpBox(
                'Validation failed. Scripts will not be written for ' + title)

    return valid_template, additional_box_val


def getSQLTableInfo(sql_table_name, cursor):
    '''Gets the values for columns: info.COLUMN_NAME, info.DATA_TYPE,
    info.IS_NULLABLE, and sy.is_identity from a specified SQL table

    :param1 sql_table_name: str
    :param2 cursor: pyodbc.cursor

    :return: List[str], List[str], List[int], List[int]
    '''

    cursor.execute("SELECT info.COLUMN_NAME, info.DATA_TYPE, info.IS_NULLABLE, sy.is_identity FROM INFORMATION_SCHEMA.COLUMNS info, sys.columns sy WHERE info.TABLE_NAME = '" +
                   sql_table_name + "' AND sy.object_id = object_id('" + sql_table_name + "') AND sy.name = info.COLUMN_NAME;")

    # Lists used to hold the values retireved from SQL script in their corresponding indexes
    sql_column_names = []
    sql_column_types = []
    column_is_nullable = []
    column_is_identity = []

    # populate row lists with values from the select script
    for row in cursor:
        sql_column_names.append(row[0])
        sql_column_types.append(row[1])
        column_is_nullable.append(row[2])
        column_is_identity.append(row[3])

    return sql_column_names, sql_column_types, column_is_nullable, column_is_identity


def connectToSQLServer():
    '''Connects to an instance of a SQL Server and allows the user to choose a
    database to work with on that instance.

    :return: List[str], pyodbc.cursor
    '''
    '''
    computer_name = str(subprocess.run(["hostname.exe"], text=True, stdout=subprocess.PIPE, input="").stdout).upper().split()[0]
    all_servers = subprocess.run(["sqlcmd", "-L"], text=True, stdout=subprocess.PIPE, input="").stdout.split()[1:]
    local_servers = []

    for server in all_servers:
        if computer_name in server:
            local_servers.append(server)
    '''
    # code id sqlcmd does not function on users laptop
    '''
    sql_server_name = ''

    while sql_server_name == '': # while user does not enter SQL Server instance
        description = "Please enter the name of the SQL Server where your database is located:"
        label = 'SQL Server name: '
        sql_server_name = gui.createTextEntryBox(
            description, label).get()
        if sql_server_name == '':
            gui.createPopUpBox(
                "Please enter a SQL server instance name.")
    '''
    '''
    description = "Please choose the name of the SQL Server where your database is located:"
    label = 'SQL Server name: '
    sql_server_name = gui.createDropDownBox(description, label, local_servers)
    '''
    # opens connection to specified SQL server and master DB to get list of all dbs on server
    dbs = pyodbc.connect('Driver={SQL Server};'
                         'Server=' + 'CHA1WS003746\\MSSQLSERVER2016' + ';'
                         'Database=' + 'master' + ';'
                         'Trusted_Connection=yes;')

    dbs_cursor = dbs.cursor()

    # executes SQL script on database connection to get list of all dbs on server
    dbs_cursor.execute(
        "SELECT name, database_id, create_date FROM sys.databases;")

    databases = []

    # populate databases with each database selected from the query
    for db in dbs_cursor:
        databases.append(db[0])

    description = "Please enter the name of the database where the table you'd like to work with is located:"
    label = 'SQL database name: '

    sql_database_name = gui.createDropDownBox(
        description, label, databases)

    # opens connection to specified SQL server and database
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=' + 'CHA1WS003746\\MSSQLSERVER2016' + ';'
                          'Database=' + sql_database_name + ';'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()

    # executes SQL script on database connection to get all tables in the database
    cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG='" +
                   sql_database_name + "' ORDER BY TABLE_NAME;")

    tables = []

    # populate tables with each table in the database
    for table in cursor:
        tables.append(table[0])

    return tables, cursor, sql_database_name


def getExcelCellToInsertInto(column, row):
    '''Gets the column and row of the excel spreadsheet that the script should be inserted into.

    :param1 column: int
    :param2 row: int

    :return: str
    '''

    # column is retrieved by finding the key of the LETTER_INDEX_DICT that the value(index) belongs to.
    excel_column = list(LETTER_INDEX_DICT.keys())[list(
        LETTER_INDEX_DICT.values()).index(column)]
    excel_row = str(row + 1)
    # excel coordinate cell that script should be inserted into
    excel_cell = excel_column + excel_row

    return excel_cell
