'''
Module of 'excel.py' that handles functions related to the validation of
Excel Workbooks to be used to write SQL scripts.
Matt Saffert
1-20-2020
'''

import excel_global
from excel_constants import *
import re
import pandas as pd


def validateWorksheetSQL(worksheet):
    '''Validates the data in the passed in worksheet based on a SQL table from an
    open SQL connection.

    :param1 worksheet: pandas.core.frame.DataFrame

    :return: bool
    '''

    valid_template = True


    tables, cursor, sql_database_name = excel_global.connectToSQLServer()
    if worksheet.loc['info'][0] == None or worksheet.loc['info'][0] not in tables:
        valid_template = False
        excel_global.createPopUpBox(
            'You have not specified a valid SQL table name in cell "A1"')
        excel_global.createPopUpBox(
            'Cannot continue SQL validation.')
        return valid_template

    if worksheet.loc['info'][1] not in TYPE_OF_SCRIPTS_AVAILABLE:
        valid_template = False
        excel_global.createPopUpBox(
            'You have not specified a valid script type in cell "B1"')

    sql_column_names, sql_column_types, column_is_nullable, column_is_identity = excel_global.getSQLTableInfo(
        worksheet.loc['info'][0], cursor)

    for i in range(len(worksheet.loc['names'])):
        if (worksheet.loc['names'][i] == None or worksheet.loc['names'][i] not in sql_column_names) and (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered a column name where one is required in cell ' + excel_global.getExcelCellToInsertInto(i, COLUMN_NAMES_ROW_INDEX))

    for i in range(len(worksheet.loc['types'])):
        type = re.sub("[\(\[].*?[\)\]]", "", str(worksheet.loc['types'][i]))
        if type not in SQL_STRING_TYPE and type not in SQL_NUMERIC_TYPE and type not in SQL_DATETIME_TYPE and type not in SQL_OTHER_TYPE:
            if (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
                valid_template = False
                excel_global.createPopUpBox(
                    'You have not entered a supported SQL type where one is required in cell ' + excel_global.getExcelCellToInsertInto(i, COLUMN_DATA_TYPE_ROW_INDEX))
        column_name = worksheet.loc['names'][i]
        if column_name in sql_column_names:
            sql_name_index = sql_column_names.index(column_name)
            if type != sql_column_types[sql_name_index]:
                valid_template = False
                excel_global.createPopUpBox(
                    'The type in your spreadsheet for ' + column_name + ', does not match the type of the column in SQL in cell ' + excel_global.getExcelCellToInsertInto(i, COLUMN_DATA_TYPE_ROW_INDEX))

    for i in range(len(worksheet.loc['include'])):
        if worksheet.loc['include'][i] != None and worksheet.loc['include'][i] != 'include':
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered an invalid string in cell ' + excel_global.getExcelCellToInsertInto(i, INCLUDE_ROW_INDEX) + '. Valid string for row 4 is "include" or leave blank')
        if worksheet.loc['info'][1] != 'delete':
            if column_is_identity[i] == 0:
                # if script type is insert, and column cannot be null then automatically select
                if column_is_nullable[i] == 'NO' and worksheet.loc['info'][1] not in ('select', 'update'):
                    if worksheet.loc['include'][i] != 'include':
                        valid_template = False
                        excel_global.createPopUpBox(
                            'You have entered an invalid string in cell ' + excel_global.getExcelCellToInsertInto(i, INCLUDE_ROW_INDEX) + '. This column must be included')
            else:  # column is identity column so cannot be updated or inserted into.
                # insert/update on identity column is NOT allowed
                if worksheet.loc['info'][1] != 'select':
                    if worksheet.loc['include'][i] == 'include':
                        valid_template = False
                        excel_global.createPopUpBox(
                            'You have entered an invalid string in cell ' + excel_global.getExcelCellToInsertInto(i, INCLUDE_ROW_INDEX) + '. This column cannot be included')

    for i in range(len(worksheet.loc['where'])):
        if worksheet.loc['where'][i] != None and worksheet.loc['where'][i] != 'where':
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered an invalid string in a cell in cell ' + excel_global.getExcelCellToInsertInto(i, WHERE_ROW_INDEX) + '. Valid string for row 5 is "where" or leave blank')

    return validateData(worksheet) and valid_template


def validateWorksheetGeneric(worksheet):
    '''Validates the data in the passed in worksheet based on a generic SQL table.

    :param1 worksheet: pandas.core.frame.DataFrame

    :return: bool
    '''

    valid_template = True

    if pd.isnull(worksheet.loc['info'][0]):
        valid_template = False
        excel_global.createPopUpBox(
            'You have not specified a SQL table name in cell "A1"')
    if worksheet.loc['info'][1] not in TYPE_OF_SCRIPTS_AVAILABLE:
        valid_template = False
        excel_global.createPopUpBox(
            'You have not specified a valid script type in cell "B1"')

    for i in range(len(worksheet.loc['names'])):
        if pd.isnull(worksheet.loc['names'][i]) and (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered a column name where one is required in cell ' + excel_global.getExcelCellToInsertInto(i, COLUMN_NAMES_ROW_INDEX))

    for i in range(len(worksheet.loc['types'])):
        type = re.sub("[\(\[].*?[\)\]]", "",
                      str(worksheet.loc['types'][i]))
        if type not in SQL_STRING_TYPE and type not in SQL_NUMERIC_TYPE and type not in SQL_DATETIME_TYPE and type not in SQL_OTHER_TYPE:
            if (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
                valid_template = False
                excel_global.createPopUpBox(
                    'You have not entered a supported SQL type where one is required in cell ' + excel_global.getExcelCellToInsertInto(i, COLUMN_DATA_TYPE_ROW_INDEX))

    for i in range(len(worksheet.loc['include'])):
        if not (pd.isnull(worksheet.loc['include'][i])) and worksheet.loc['include'][i] != 'include':
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered an invalid string in cell ' + excel_global.getExcelCellToInsertInto(i, INCLUDE_ROW_INDEX) + '. Valid string for row 4 is "include" or leave blank')

    for i in range(len(worksheet.loc['where'])):
        if not (pd.isnull(worksheet.loc['where'][i])) and worksheet.loc['where'][i] != 'where':
            valid_template = False
            excel_global.createPopUpBox(
                'You have not entered an invalid string in a cell in cell ' + excel_global.getExcelCellToInsertInto(i, WHERE_ROW_INDEX) + '. Valid string for row 5 is "where" or leave blank')

    return validateData(worksheet) and valid_template


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
                excel_global.createPopUpBox(
                    'You have not entered a value in cell ' + excel_global.getExcelCellToInsertInto(i, row) + ' where one is required')
    blank_last_row = True
    for i in range(len(worksheet.iloc[len(worksheet) - 1])):
        if not (pd.isnull(worksheet.iloc[len(worksheet) - 1][i])):
            blank_last_row = False
    if not blank_last_row:
        for i in range(len(worksheet.iloc[len(worksheet) - 1])):
            if pd.isnull(worksheet.iloc[len(worksheet) - 1][i]) and (worksheet.loc['include'][i] == 'include' or worksheet.loc['where'][i] == 'where'):
                valid_template = False
                excel_global.createPopUpBox(
                    'You have not entered a value in cell ' + excel_global.getExcelCellToInsertInto(i, len(worksheet) - 1) + ' where one is required')

    return valid_template


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
    write_script_for = excel_global.createYesNoBox(
        description, yes, no)

    valid_template = True
    if validate_with_sql == 'Generic':
        if write_script_for == yes:  # if the user says to write scripts for this sheet
            valid_template = validateWorksheetGeneric(worksheet) and valid_template
        else:
            valid_template = False
            excel_global.createPopUpBox(
                'Validation failed. Scripts will not be written for ' + title)

    elif validate_with_sql == 'SQL':
        if write_script_for == yes:  # if the user says to write scripts for this sheet
            valid_template = validateWorksheetSQL(worksheet) and valid_template
        else:
            valid_template = False
            excel_global.createPopUpBox(
                'Validation failed. Scripts will not be written for ' + title)

    return valid_template


def validate(workbook):
    '''Cycles through worksheets in a workbook checking if they're valid.
    Alerts the user if any or all worksheets are invalid.

    :param1 workbook: dict
    '''

    validate_with_sql = excel_global.createYesNoBox(  # tkinter dialog box that asks user if they want to connect to a SQL database to validate spreadsheet
        'Would you like to validate Workbook with SQL table or generic validation?', 'SQL', 'Generic')

    any_changes = False # False if all spreadsheets fail validation
    all_sheets_okay = True # True if all spreadsheets pass validation

    for worksheet in workbook:
        # check if worksheet is is valid and if user wants to write scripts for them
        valid_template = validWorksheet(
            workbook[worksheet], validate_with_sql)
        all_sheets_okay = valid_template # True if spreadsheet passes validation
        if valid_template:  # only write to Excel if the Excel spreadsheet is a valid format
            output_string = "VALID. This worksheet will function properly with the 'Write SQL script' mode of this program."
            excel_global.createPopUpBox(
                output_string)  # tkinter dialog box
            any_changes = True  # changes were made and need to be saved

    if any_changes and not all_sheets_okay:  # some but not all spreadsheets in workbook pass validation
        output_string = "CAUTION. Care must be taken building scripts with this workbook because not all sheets are in a valid form."
        excel_global.createPopUpBox(output_string)
