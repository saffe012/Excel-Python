'''
Module of 'excel.py' that handles functions related to generating and writing SQL
scripts to an Excel spreadsheet.
Matt Saffert
1-9-2020
'''

from excel_constants import *
import re
import tkinter
import excel_global
from tkinter import filedialog as tkFileDialog
import pandas as pd
import global_gui as gui


def writeMode():
    '''Starts the writing script mode of the application.

    :return: NONE
    '''

    # displays how Excel spreadsheet should be laid out
    gui.createPopUpBox(TEMPLATE_DESCRIPTION, "600x500")  # tkinter dialog box

    output_string = "Choose the Excel workbook you'd like to make scripts for."
    workbook = gui.openExcelFile(output_string)

    validate_with_sql, additional_box_val = gui.createTwoChoiceBox(
        'Would you like to validate Workbook with SQL table or generic validation?', 'SQL', 'Generic')

    write_to_sql = 'SQL'
    write_to_excel = 'Excel'
    description = 'Would you like to write the sql scripts to a ".sql" file or to an Excel spreadsheet?'
    write_to, additional_box_val = gui.createTwoChoiceBox(  # write scripts to new SQL or Excel file
        description, write_to_sql, write_to_excel)

    if write_to == 'SQL':
        save_file = writeToSQL(workbook, validate_with_sql)
    elif write_to == 'Excel':
        save_file = writeToExcel(workbook, validate_with_sql)

    if save_file == '':  # no scripts were written because there were no valid worksheets
        output_string = "No files were changed. Closing program."
        gui.createPopUpBox(
            output_string)  # tkinter dialog box


def isValueTypeString(type):
    '''Checks the SQL type of the column of data in the spreadsheet based on the type
    row in the excel spreadsheet. Returns true id type needs parenthesis around it
    in the script

    :param1 type: str

    :return: bool
    '''

    # ex. 'varchar(200)' -> 'varchar'
    # regular expression that strips parenthesis off end of type
    type = re.sub("[\(\[].*?[\)\]]", "", str(type))

    # decides whether the value of this type needs parenthesis around it in script
    if (type in SQL_STRING_TYPE) or (type in SQL_DATETIME_TYPE) or (type in SQL_OTHER_TYPE):
        return True
    elif type == 'bit':  # bit can be represented by both 1/0 integers, or 'True'/'False' strings. This program uses strings
        return True
    else:
        return False


def shouldInclude(value):
    '''Checks whether a column of data should be included in the SQL script based on
    the include row of the excel spreadsheet.

    :param1 value: str

    :return: bool
    '''

    include = str(value)
    if include == 'include':
        return True
    return False


def includeInWhereClause(value):
    '''Checks whether a column of data should be included in the where clause of the
    generated SQL script based on the where row of the excel spreadsheet.

    :param1 value: str

    :return: bool
    '''

    wheres = str(value)
    if wheres == 'where':
        return True
    return False


def writeScripts(worksheet):
    '''Checks the desired type of SQL script to be generated and calls the corresponding
    function the generate scripts.

    :param1 worksheet: pandas.core.frame.DataFrame

    :return: dict
    '''

    scripts = {}
    script_type = worksheet.loc['info'][SCRIPT_TYPE]

    if script_type == 'insert':
        scripts = createInsertScripts(worksheet)
    elif script_type == 'update':
        scripts = createUpdateScripts(worksheet)
    elif script_type == 'delete':
        scripts = createDeleteScripts(worksheet)
    elif script_type == 'select':
        scripts = createSelectScripts(worksheet)

    return scripts


def createColumnClause(worksheet, statement):
    '''Helper function to createInsertScripts() that creates the list of column
    names to insert/select for the SQL script for each row of the excel spreadsheet

    :param1 worksheet: pandas.core.frame.DataFrame
    :param2 statement: str

    :return: str
    '''

    # concatenates each included value of each column to the return string
    for i in range(len(worksheet.loc['names']) - 1):
        if shouldInclude(worksheet.loc['include'][i]):
            statement = ''.join(
                [statement, (str(worksheet.loc['names'][i]) + ', ')])
    # checks last column
    if shouldInclude(worksheet.loc['include'][len(worksheet.loc['names']) - 1]):
        statement = ''.join(
            [statement, (str(worksheet.loc['names'][len(worksheet.loc['names']) - 1]))])
    else:
        # if last column should not be included drop last space and comma from string
        statement = statement[:-2]

    return statement


def createValuesClause(worksheet, statement, row):
    '''Helper function to createInsertScripts() that creates the VALUES clause of
    the SQL script for each row of the excel spreadsheet

    :param1 worksheet: pandas.core.frame.DataFrame
    :param2 statement: str
    :param3 row: int

    :return: str
    '''

    # concatenates each included value of each column to the return string
    for cell in range(len(worksheet.iloc[row]) - 1):
        if shouldInclude(worksheet.loc['include'][cell]):
            string = isValueTypeString(worksheet.loc['types'][cell])
            if string:  # add quotes
                statement = ''.join(
                    [statement, ("'" + str(worksheet.iloc[row][cell]) + "', ")])
            else:
                statement = ''.join(
                    [statement, (str(worksheet.iloc[row][cell]) + ", ")])
    # checks last column
    if shouldInclude(worksheet.loc['include'][len(worksheet.iloc[row]) - 1]):
        string = isValueTypeString(
            worksheet.loc['types'][len(worksheet.iloc[row]) - 1])
        if string:  # add quotes
            statement = ''.join(
                [statement, ("'" + str(worksheet.iloc[row][len(worksheet.iloc[row]) - 1]) + "');")])
        else:
            statement = ''.join(
                [statement, (str(worksheet.iloc[row][len(worksheet.iloc[row]) - 1]) + ");")])
    else:
        # if last column should not be included drop last space and comma from string
        statement = statement[:-2] + ');'

    return statement


def createInsertScripts(worksheet):
    '''Creates the insert scripts based on the data provided in the Excel spreadsheet.

    :param1 table_name: pandas.core.frame.DataFrame
    :param2 worksheet: tuple

    :return: dict
    '''

    script_dict = {}  # {cell: script}. ex. {'G7': 'INSERT INTO... ;'}
    pre_statement = 'INSERT INTO ' + \
        worksheet.loc['info'][TABLE_NAME] + ' ('

    insert_statement = createColumnClause(
        worksheet, pre_statement) + ') VALUES ('

    # creates script for each row of data in the Excel table
    for row in range(START_OF_DATA_ROWS_INDEX, len(worksheet) - 1):
        values_statement = createValuesClause(
            worksheet, insert_statement, row)

        excel_cell = excel_global.getExcelCellToInsertInto(
            len(worksheet.iloc[row]), row)
        script_dict[excel_cell] = values_statement

    all_none = True
    # look at the last row. Row may be blank and generate. None values. so don't write scripts
    for i in range(len(worksheet.iloc[len(worksheet) - 1])):
        # if there is one value that isn't None, write scripts
        if worksheet.iloc[len(worksheet) - 1][i] != None:
            all_none = False
    if not all_none:
        values_statement = createValuesClause(
            worksheet, insert_statement, len(worksheet) - 1)
        excel_cell = excel_global.getExcelCellToInsertInto(
            len(worksheet.iloc[len(worksheet) - 1]), len(worksheet) - 1)
        script_dict[excel_cell] = values_statement

    return script_dict


def createUpdateClause(worksheet, statement, row):
    '''Helper function to createUpdateScripts() that creates the UPDATE clause of the
    SQL script for each row of the excel spreadsheet

    :param1 worksheet: pandas.core.frame.DataFrame
    :param5 statement: str
    :param6 row: int

    :return: str
    '''

    # concatenates each included value of each column to the return string
    for cell in range(len(worksheet.iloc[row]) - 1):
        if shouldInclude(worksheet.loc['include'][cell]):
            statement = ''.join(
                [statement, (str(worksheet.loc['names'][cell]) + ' = ')])
            string = isValueTypeString(worksheet.loc['types'][cell])
            if string:  # add quotes
                statement = ''.join(
                    [statement, ("'" + str(worksheet.iloc[row][cell]) + "', ")])
            else:
                statement = ''.join(
                    [statement, (str(worksheet.iloc[row][cell]) + ", ")])
    # checks last column
    if shouldInclude(worksheet.loc['include'][len(worksheet.iloc[row]) - 1]):
        statement = ''.join(
            [statement, (str(worksheet.loc['names'][len(worksheet.loc['names']) - 1]) + ' = ')])
        string = isValueTypeString(
            worksheet.loc['types'][len(worksheet.loc['names']) - 1])
        if string:  # add quotes
            statement = ''.join(
                [statement, ("'" + str(worksheet.iloc[row][len(worksheet.iloc[row]) - 1]) + "' WHERE ")])
        else:
            statement = ''.join(
                [statement, (str(worksheet.iloc[row][len(worksheet.iloc[row]) - 1]) + " WHERE ")])
    else:
        # if last column should not be included drop last space and comma from string
        statement = statement[:-2] + ' WHERE '

    return statement


def createWhereClause(worksheet, statement, row):
    '''Helper function to createUpdateScripts() that creates the WHERE clause of the
    SQL script for each row of the excel spreadsheet

    :param1 worksheet: pandas.core.frame.DataFrame
    :param5 statement: str
    :param6 row: int

    :return: str
    '''

    # concatenates each where value of each column to the return string
    for i in range(len(worksheet.iloc[row]) - 1):
        if includeInWhereClause(worksheet.loc['where'][i]):
            statement = ''.join(
                [statement, (str(worksheet.loc['names'][i]) + ' = ')])
            string = isValueTypeString(worksheet.loc['types'][i])
            if string:  # add quotes
                statement = ''.join(
                    [statement, ("'" + str(worksheet.iloc[row][i]) + "'  AND  ")])
            else:
                statement = ''.join(
                    [statement, (str(worksheet.iloc[row][i]) + "  AND  ")])
    # checks last column
    if includeInWhereClause(worksheet.loc['where'][len(worksheet.loc['names']) - 1]):
        statement = ''.join(
            [statement, (str(worksheet.loc['names'][len(worksheet.loc['names']) - 1]) + ' = ')])
        string = isValueTypeString(
            worksheet.loc['types'][len(worksheet.loc['names']) - 1])
        if string:  # add quotes
            statement = ''.join(
                [statement, ("'" + str(worksheet.iloc[row][len(worksheet.iloc[row]) - 1]) + "';")])
        else:
            statement = ''.join(
                [statement, (str(worksheet.iloc[row][len(worksheet.iloc[row]) - 1]) + ";")])
    else:
        # if last column is not in where clause drop 'AND' statement or drop ' WHERE' statement
        statement = statement[:-7] + ';'

    return statement


def createUpdateScripts(worksheet):
    '''Creates the update scripts based on the data provided in the Excel spreadsheet.

    :param1 worksheet: pandas.core.frame.DataFrame

    :return: dict
    '''

    script_dict = {}  # {cell: script}. ex. {'G7': 'UPDATE... ;'}
    pre_statement = 'UPDATE ' + \
        worksheet.loc['info'][TABLE_NAME] + ' SET '

    # creates script for each row of data in the Excel table
    for row in range(START_OF_DATA_ROWS_INDEX, len(worksheet)):
        update_statement = createUpdateClause(
            worksheet, pre_statement, row)

        where_statement = createWhereClause(
            worksheet, update_statement, row)

        excel_cell = excel_global.getExcelCellToInsertInto(
            len(worksheet.iloc[row]), row)
        script_dict[excel_cell] = where_statement

    return script_dict


def createDeleteScripts(worksheet):
    '''Creates the delete scripts based on the data provided in the Excel spreadsheet.

    :param1 worksheet: pandas.core.frame.DataFrame

    :return: dict
    '''

    script_dict = {}  # {cell: script}. ex. {'G7': 'DELETE... ;'}
    pre_statement = 'DELETE FROM ' + \
        worksheet.loc['info'][TABLE_NAME] + ' WHERE '

    # creates script for each row of data in the Excel table
    for row in range(START_OF_DATA_ROWS_INDEX, len(worksheet)):
        where_statement = createWhereClause(
            worksheet, pre_statement, row)

        excel_cell = excel_global.getExcelCellToInsertInto(
            len(worksheet.iloc[row]), row)
        script_dict[excel_cell] = where_statement

    return script_dict


def createSelectScripts(worksheet):
    '''Creates the select scripts based on the data provided in the Excel spreadsheet.

    :param1 worksheet: pandas.core.frame.DataFrame

    :return: dict
    '''

    script_dict = {}  # {cell: script}. ex. {'G7': 'INSERT INTO... ;'}
    pre_statement = 'SELECT ('

    select_statement = createColumnClause(
        worksheet, pre_statement) + ') FROM ' + worksheet.loc['info'][TABLE_NAME] + ' WHERE '

    # creates script for each row of data in the Excel table
    for row in range(START_OF_DATA_ROWS_INDEX, len(worksheet)):
        where_statement = createWhereClause(
            worksheet, select_statement, row)

        excel_cell = excel_global.getExcelCellToInsertInto(
            len(worksheet.iloc[row]), row)
        script_dict[excel_cell] = where_statement

    return script_dict


def writeToExcel(workbook, validate_with_sql):
    '''Iterates through each worksheet in the imported workbook, creates
    scripts for each worksheet, and writes the scripts to a new workbook. Returns
    True if scripts were generated and need to be saved, otherwise False

    :param1 workbook: dict
    :param2 validate_with_sql: str

    :return: bool
    '''

    any_changes = ''
    valid_template = True

    for worksheet in workbook:
        valid_template, additional_box_val = excel_global.validWorksheet(
            workbook[worksheet], validate_with_sql, worksheet)

        if valid_template:  # only write to Excel if the Excel spreadsheet is a valid format

            # returns dict containing excel cell coordinates as key and script to writeas value
            scripts = writeScripts(workbook[worksheet])

            any_changes = 'Excel'  # changes were made and need to be saved
            # writes script to worksheet
            df_scripts = ['', 'Scripts', '', '', '']
            for cell, script in scripts.items():
                df_scripts.append(script)
            workbook[worksheet]['scripts'] = df_scripts
    #
    if valid_template:
        gui.saveToExcel(workbook)

    return any_changes


def saveToSQL(text_file):
    '''Saves the string to a SQL file.

    :param1 text_file: str
    '''

    file = tkinter.Tk()
    # opens file explorer so user can choose file to write to
    file.filename = tkFileDialog.asksaveasfilename(
        initialdir="C:/", title="Select/create file to save/write to", defaultextension=".sql")
    f = open(file.filename, 'w')
    f.write(text_file)
    f.close()
    file.destroy()

    output_string = "Scripts saved to: '" + \
        str(file.filename) + "'"
    gui.createPopUpBox(
        output_string)  # tkinter dialog box


def writeToSQL(workbook, validate_with_sql):
    '''Iterates through each worksheet in the imported workbook, creates
    scripts for each worksheet, and writes the scripts to a SQL file. Returns
    True if scripts were generated and have been saved, otherwise False

    :param1 workbook: dict
    :param2 validate_with_sql: str

    :return: bool
    '''

    any_changes = ''
    text_file = ''
    valid_template = True

    for worksheet in workbook:
        valid_template, additional_box_val = excel_global.validWorksheet(
            workbook[worksheet], validate_with_sql, worksheet)
        if valid_template:  # only write to Excel if the Excel spreadsheet is a valid format

            # returns dict containing excel cell coordinates as key and script to write as value
            scripts = writeScripts(workbook[worksheet])

            any_changes = 'SQL'

            for cell, script in scripts.items():
                text_file += script + '\n'
    #
    if valid_template:
        saveToSQL(text_file)

    return any_changes
