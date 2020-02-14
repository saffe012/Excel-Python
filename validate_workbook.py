'''
Module of 'excel.py' that handles functions related to the validation of
Excel Workbooks to be used to write SQL scripts.
Matt Saffert
1-20-2020
'''

import excel_global
import constants as cons
import re


def validateWorksheetSQL(table):
    '''Validates the data in the passed in table based on a SQL table from an
    open SQL connection.

    :param1 table: tuple

    :return: bool
    '''

    valid_template = True
    column_info = table[cons.INFO_ROW]
    column_names = table[cons.COLUMN_NAMES_ROW_INDEX]
    column_types = table[cons.COLUMN_DATA_TYPE_ROW_INDEX]
    column_includes = table[cons.INCLUDE_ROW_INDEX]
    column_where = table[cons.WHERE_ROW_INDEX]

    tables, cursor, sql_database_name = excel_global.connectToSQLServer()
    if column_info[0].value == None or column_info[0].value not in tables:
        valid_template = False
        excel_global.createPopUpBox(
            'You have not specified a valid SQL table name in cell "A1"')
        excel_global.createPopUpBox(
            'Cannot continue SQL validation.')
        return valid_template

    if column_info[1].value not in cons.TYPE_OF_SCRIPTS_AVAILABLE:
        valid_template = False
        excel_global.createPopUpBox(
            'You have not specified a valid script type in cell "B1"')

    sql_column_names, sql_column_types, column_is_nullable, column_is_identity = excel_global.getSQLTableInfo(
        column_info[0].value, cursor)

    for i in range(len(column_names)):
        if (column_names[i].value == None or column_names[i].value not in sql_column_names) and (column_includes[i].value == 'include' or column_where[i].value == 'where'):
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered a column name where one is required in cell ' + excel_global.getExcelCellToInsertInto(i, cons.COLUMN_NAMES_ROW_INDEX))

    for i in range(len(column_types)):
        type = re.sub("[\(\[].*?[\)\]]", "", str(column_types[i].value))
        if type not in cons.SQL_STRING_TYPE and type not in cons.SQL_NUMERIC_TYPE and type not in cons.SQL_DATETIME_TYPE and type not in cons.SQL_OTHER_TYPE:
            if (column_includes[i].value == 'include' or column_where[i].value == 'where'):
                valid_template = False
                excel_global.createPopUpBox(
                    'You have not entered a supported SQL type where one is required in cell ' + excel_global.getExcelCellToInsertInto(i, cons.COLUMN_DATA_TYPE_ROW_INDEX))
        column_name = column_names[i].value
        if column_name in sql_column_names:
            sql_name_index = sql_column_names.index(column_name)
            if type != sql_column_types[sql_name_index]:
                valid_template = False
                excel_global.createPopUpBox(
                    'The type in your spreadsheet for ' + column_name + ', does not match the type of the column in SQL in cell ' + excel_global.getExcelCellToInsertInto(i, cons.COLUMN_DATA_TYPE_ROW_INDEX))

    for i in range(len(column_includes)):
        if column_includes[i].value != None and column_includes[i].value != 'include':
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered an invalid string in cell ' + excel_global.getExcelCellToInsertInto(i, cons.INCLUDE_ROW_INDEX) + '. Valid string for row 4 is "include" or leave blank')
        if column_info[1].value != 'delete':
            if column_is_identity[i] == 0:
                # if script type is insert, and column cannot be null then automatically select
                if column_is_nullable[i] == 'NO' and column_info[1].value not in ('select', 'update'):
                    if column_includes[i].value != 'include':
                        valid_template = False
                        excel_global.createPopUpBox(
                            'You have entered an invalid string in cell ' + excel_global.getExcelCellToInsertInto(i, cons.INCLUDE_ROW_INDEX) + '. This column must be included')
            else:  # column is identity column so cannot be updated or inserted into.
                # insert/update on identity column is NOT allowed
                if column_info[1].value != 'select':
                    if column_includes[i].value == 'include':
                        valid_template = False
                        excel_global.createPopUpBox(
                            'You have entered an invalid string in cell ' + excel_global.getExcelCellToInsertInto(i, cons.INCLUDE_ROW_INDEX) + '. This column cannot be included')

    for i in range(len(column_where)):
        if column_where[i].value != None and column_where[i].value != 'where':
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered an invalid string in a cell in cell ' + excel_global.getExcelCellToInsertInto(i, cons.WHERE_ROW_INDEX) + '. Valid string for row 5 is "where" or leave blank')

    return validateData(table) and valid_template


def validateWorksheetGeneric(table):
    '''Validates the data in the passed in table based on a generic SQL table.

    :param1 table: tuple

    :return: bool
    '''

    valid_template = True
    column_info = table[cons.INFO_ROW]
    column_names = table[cons.COLUMN_NAMES_ROW_INDEX]
    column_types = table[cons.COLUMN_DATA_TYPE_ROW_INDEX]
    column_includes = table[cons.INCLUDE_ROW_INDEX]
    column_where = table[cons.WHERE_ROW_INDEX]

    if column_info[0].value == None:
        valid_template = False
        excel_global.createPopUpBox(
            'You have not specified a SQL table name in cell "A1"')
    if column_info[1].value not in cons.TYPE_OF_SCRIPTS_AVAILABLE:
        valid_template = False
        excel_global.createPopUpBox(
            'You have not specified a valid script type in cell "B1"')

    for i in range(len(column_names)):
        if column_names[i].value == None and (column_includes[i].value == 'include' or column_where[i].value == 'where'):
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered a column name where one is required in cell ' + excel_global.getExcelCellToInsertInto(i, cons.COLUMN_NAMES_ROW_INDEX))

    for i in range(len(column_types)):
        type = re.sub("[\(\[].*?[\)\]]", "",
                      str(column_types[i].value))
        if type not in cons.SQL_STRING_TYPE and type not in cons.SQL_NUMERIC_TYPE and type not in cons.SQL_DATETIME_TYPE and type not in cons.SQL_OTHER_TYPE:
            if (column_includes[i].value == 'include' or column_where[i].value == 'where'):
                valid_template = False
                excel_global.createPopUpBox(
                    'You have not entered a supported SQL type where one is required in cell ' + excel_global.getExcelCellToInsertInto(i, cons.COLUMN_DATA_TYPE_ROW_INDEX))

    for i in range(len(column_includes)):
        if column_includes[i].value != None and column_includes[i].value != 'include':
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered an invalid string in cell ' + excel_global.getExcelCellToInsertInto(i, cons.INCLUDE_ROW_INDEX) + '. Valid string for row 4 is "include" or leave blank')

    for i in range(len(column_where)):
        if column_where[i].value != None and column_where[i].value != 'where':
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered an invalid string in a cell in cell ' + excel_global.getExcelCellToInsertInto(i, cons.WHERE_ROW_INDEX) + '. Valid string for row 5 is "where" or leave blank')

    return validateData(table) and valid_template


def validateData(table):
    '''Validates the data in the passed in table. The data comes from the 6th
    row and on in an Excel spreadsheet

    :param1 table: tuple

    :return: bool
    '''

    valid_template = True
    column_includes = table[cons.INCLUDE_ROW_INDEX]
    column_where = table[cons.WHERE_ROW_INDEX]

    for row in range(cons.START_OF_DATA_ROWS_INDEX, len(table) - 1):
        for i in range(len(table[row])):
            if table[row][i].value == None and (column_includes[i].value == 'include' or column_where[i].value == 'where'):
                valid_template = False
                excel_global.createPopUpBox(
                    'You have not entered a value in cell ' + excel_global.getExcelCellToInsertInto(i, row) + ' where one is required')
    blank_last_row = True
    for i in range(len(table[len(table) - 1])):
        if table[len(table) - 1][i].value != None:
            blank_last_row = False
    if not blank_last_row:
        for i in range(len(table[len(table) - 1])):
            if table[len(table) - 1][i].value == None and (column_includes[i].value == 'include' or column_where[i].value == 'where'):
                valid_template = False
                excel_global.createPopUpBox(
                    'You have not entered a value in cell ' + excel_global.getExcelCellToInsertInto(i, len(table) - 1) + ' where one is required')

    return valid_template


def validWorksheet(worksheet, validate_with_sql):
    '''Calls the correct function to validate the passed worksheet based on
    whether a user wants to connect to SQL or not.

    :param1 worksheet: openpyxl.worksheet.worksheet.Worksheet
    :param2 validate_with_sql: str

    :return: bool
    '''

    # TODO:
    # Finish valid_template
    # add seperate check template function that will connect to database

    table = tuple(worksheet.rows)
    column_info = table[cons.INFO_ROW]
    column_names = table[cons.COLUMN_NAMES_ROW_INDEX]
    column_types = table[cons.COLUMN_DATA_TYPE_ROW_INDEX]
    column_includes = table[cons.INCLUDE_ROW_INDEX]
    column_where = table[cons.WHERE_ROW_INDEX]

    description = "Would you like to validate/create scripts for " + \
        worksheet.title + " worksheet?"
    yes = "Yes"
    no = "No"
    write_script_for = excel_global.createYesNoBox(
        description, yes, no)

    valid_template = True
    if validate_with_sql == 'Generic':
        if write_script_for == yes:  # if the user says to write scripts for this sheet
            valid_template = validateWorksheetGeneric(table) and valid_template
        else:
            valid_template = False
            excel_global.createPopUpBox(
                'Validation failed. Scripts will not be written for ' + worksheet.title)

    elif validate_with_sql == 'SQL':
        if write_script_for == yes:  # if the user says to write scripts for this sheet
            valid_template = validateWorksheetSQL(table) and valid_template
        else:
            valid_template = False
            excel_global.createPopUpBox(
                'Validation failed. Scripts will not be written for ' + worksheet.title)

    return valid_template
