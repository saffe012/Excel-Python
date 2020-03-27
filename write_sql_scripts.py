'''
Module of 'excel.py' that handles functions related to generating and writing SQL
scripts to an Excel spreadsheet.
Matt Saffert
1-9-2020
'''

import constants as cons
import re
import tkinter
import excel_global
import validate_workbook as validate
from tkinter import filedialog as tkFileDialog


def displayExcelFormatInstructions():
    '''Creates a tkinter pop-up box that explains the formatting of the excel
    spreadsheet that will be read into the program

    :return: NONE
    '''

    root = tkinter.Tk()
    excel_global.addQuitMenuButton(root)
    root.title('Excel Python')
    root.geometry("600x500")
    w = tkinter.Label(root, text='Please make sure the excel spreadsheet that '
                      'will be read was made with the tool and/or is formatted '
                      'correctly:\nRow 1: col1: SQL tablename col2: script type\nRow 2: SQL column '
                      'names\nRow 3: SQL data types\nRow 4: put "include" in '
                      'cells you want to be inserted/updated\nRow 5: put "where" '
                      'in cells you want to be included in delete/update where '
                      'clause. (For inserts, leave blank)\nRow 6: Start of data')
    w.pack()
    w.place(relx=0.5, rely=0.2, anchor='center')
    button = tkinter.Button(root, text='Ok', width=25, command=root.destroy).place(
        relx=0.5, rely=0.5, anchor='center')
    root.mainloop()


def getTypeOfScript(worksheet):
    '''Gets the type of scripts as labeled in excel sheet

    :param1 worksheet: openpyxl.worksheet.worksheet.Worksheet

    :return: str
    '''

    all_rows = tuple(worksheet.rows)
    info_row = all_rows[cons.INFO_ROW]
    script_type = info_row[1].value
    return script_type


def getTableName(worksheet):
    '''Gets the SQL table name from scripts as labeled in excel sheet

    :param1 worksheet: openpyxl.worksheet.worksheet.Worksheet

    :return: str
    '''

    all_rows = tuple(worksheet.rows)
    info_row = all_rows[cons.INFO_ROW]
    table_name = info_row[0].value
    return table_name


def isValueTypeString(types, column):
    '''Checks the SQL type of the column of data in the spreadsheet based on the type
    row in the excel spreadsheet. Returns true id type needs parenthesis around it
    in the script

    :param1 types: list
    :param2 column: int

    :return: bool
    '''

    # ex. 'varchar(200)' -> 'varchar'
    # regular expression that strips parenthesis off end of type
    type = re.sub("[\(\[].*?[\)\]]", "", str(types[column].value))

    # decides whether the value of this type needs parenthesis around it in script
    if (type in cons.SQL_STRING_TYPE) or (type in cons.SQL_DATETIME_TYPE) or (type in cons.SQL_OTHER_TYPE):
        return True
    elif type == 'bit':  # bit can be represented by both 1/0 integers, or 'True'/'False' strings. This program uses strings
        return True
    else:
        return False


def shouldInclude(includes, column):
    '''Checks whether a column of data should be included in the SQL script based on
    the include row of the excel spreadsheet.

    :param1 includes: list
    :param2 column: int

    :return: bool
    '''

    include = str(includes[column].value)
    if include == 'include':
        return True
    return False


def includeInWhereClause(where, column):
    '''Checks whether a column of data should be included in the where clause of the
    generated SQL script based on the where row of the excel spreadsheet.

    :param1 where: list
    :param2 column: int

    :return: bool
    '''

    wheres = str(where[column].value)
    if wheres == 'where':
        return True
    return False


def writeScripts(table, script_type, table_name):
    '''Checks the desired type of SQL script to be generated and calls the corresponding
    function the generate scripts.

    :param1 table: tuple
    :param2 script_type: str
    :param3 table_name: str

    :return: dict
    '''

    scripts = {}

    if script_type == 'insert':
        scripts = createInsertScripts(table_name, table)
    elif script_type == 'update':
        scripts = createUpdateScripts(table_name, table)
    elif script_type == 'delete':
        scripts = createDeleteScripts(table_name, table)
    elif script_type == 'select':
        scripts = createSelectScripts(table_name, table)

    return scripts


def createColumnClause(column_names, column_includes, statement):
    '''Helper function to createInsertScripts() that creates the list of column
    names to insert/select for the SQL script for each row of the excel spreadsheet

    :param1 column_names: list
    :param2 column_includes: list
    :param3 statement: str

    :return: str
    '''

    # concatenates each included value of each column to the return string
    for i in range(len(column_names) - 1):
        if shouldInclude(column_includes, i):
            statement = ''.join(
                [statement, (str(column_names[i].value) + ', ')])
    if shouldInclude(column_includes, len(column_names) - 1):  # checks last column
        statement = ''.join(
            [statement, (str(column_names[len(column_names) - 1].value))])
    else:
        # if last column should not be included drop last space and comma from string
        statement = statement[:-2]

    return statement


def createValuesClause(table, column_types, column_includes, statement, row):
    '''Helper function to createInsertScripts() that creates the VALUES clause of
    the SQL script for each row of the excel spreadsheet

    :param1 table: tuple
    :param2 column_types: list
    :param3 column_includes: list
    :param4 statement: str
    :param5 row: int

    :return: str
    '''

    # concatenates each included value of each column to the return string
    for cell in range(len(table[row]) - 1):
        if shouldInclude(column_includes, cell):
            string = isValueTypeString(column_types, cell)
            if string:  # add quotes
                statement = ''.join(
                    [statement, ("'" + str(table[row][cell].value) + "', ")])
            else:
                statement = ''.join(
                    [statement, (str(table[row][cell].value) + ", ")])
    # checks last column
    if shouldInclude(column_includes, len(table[row]) - 1):
        string = isValueTypeString(column_types, len(table[row]) - 1)
        if string:  # add quotes
            statement = ''.join(
                [statement, ("'" + str(table[row][len(table[row]) - 1].value) + "');")])
        else:
            statement = ''.join(
                [statement, (str(table[row][len(table[row]) - 1].value) + ");")])
    else:
        # if last column should not be included drop last space and comma from string
        statement = statement[:-2] + ');'

    return statement


def createInsertScripts(table_name, table):
    '''Creates the insert scripts based on the data provided in the Excel spreadsheet.

    :param1 table_name: str
    :param2 table: tuple

    :return: dict
    '''

    script_dict = {}  # {cell: script}. ex. {'G7': 'INSERT INTO... ;'}
    column_names = table[cons.COLUMN_NAMES_ROW_INDEX]
    column_types = table[cons.COLUMN_DATA_TYPE_ROW_INDEX]
    column_includes = table[cons.INCLUDE_ROW_INDEX]

    pre_statement = 'INSERT INTO ' + table_name + ' ('

    insert_statement = createColumnClause(
        column_names, column_includes, pre_statement) + ') VALUES ('

    # creates script for each row of data in the Excel table
    for row in range(cons.START_OF_DATA_ROWS_INDEX, len(table) - 1):
        values_statement = createValuesClause(
            table, column_types, column_includes, insert_statement, row)

        excel_cell = excel_global.getExcelCellToInsertInto(
            len(table[row]), row)
        script_dict[excel_cell] = values_statement

    all_none = True
    # look at the last row. Row may be blank and generate. None values. so don't write scripts
    for i in range(len(table[len(table) - 1])):
        # if there is one value that isn't None, write scripts
        if table[len(table) - 1][i].value != None:
            all_none = False
    if not all_none:
        values_statement = createValuesClause(
            table, column_types, column_includes, insert_statement, len(table) - 1)
        excel_cell = excel_global.getExcelCellToInsertInto(
            len(table[len(table) - 1]), len(table) - 1)
        script_dict[excel_cell] = values_statement

    return script_dict


def createUpdateClause(table, column_names, column_types, column_includes, statement, row):
    '''Helper function to createUpdateScripts() that creates the UPDATE clause of the
    SQL script for each row of the excel spreadsheet

    :param1 table: tuple
    :param2 column_names: list
    :param3 column_types: list
    :param4 column_includes: list
    :param5 statement: str
    :param6 row: int

    :return: str
    '''

    # concatenates each included value of each column to the return string
    for cell in range(len(table[row]) - 1):
        if shouldInclude(column_includes, cell):
            statement = ''.join(
                [statement, (str(column_names[cell].value) + ' = ')])
            string = isValueTypeString(column_types, cell)
            if string:  # add quotes
                statement = ''.join(
                    [statement, ("'" + str(table[row][cell].value) + "', ")])
            else:
                statement = ''.join(
                    [statement, (str(table[row][cell].value) + ", ")])
    # checks last column
    if shouldInclude(column_includes, len(table[row]) - 1):
        statement = ''.join(
            [statement, (str(column_names[len(column_names) - 1].value) + ' = ')])
        string = isValueTypeString(column_types, len(column_names) - 1)
        if string:  # add quotes
            statement = ''.join(
                [statement, ("'" + str(table[row][len(table[row]) - 1].value) + "' WHERE ")])
        else:
            statement = ''.join(
                [statement, (str(table[row][len(table[row]) - 1].value) + " WHERE ")])
    else:
        # if last column should not be included drop last space and comma from string
        statement = statement[:-2] + ' WHERE '

    return statement


def createWhereClause(table, column_names, column_types, column_where, statement, row):
    '''Helper function to createUpdateScripts() that creates the WHERE clause of the
    SQL script for each row of the excel spreadsheet

    :param1 table: tuple
    :param2 column_names: list
    :param3 column_types: list
    :param4 column_where: list
    :param5 statement: str
    :param6 row: int

    :return: str
    '''

    # concatenates each where value of each column to the return string
    for i in range(len(table[row]) - 1):
        if includeInWhereClause(column_where, i):
            statement = ''.join(
                [statement, (str(column_names[i].value) + ' = ')])
            string = isValueTypeString(column_types, i)
            if string:  # add quotes
                statement = ''.join(
                    [statement, ("'" + str(table[row][i].value) + "'  AND  ")])
            else:
                statement = ''.join(
                    [statement, (str(table[row][i].value) + "  AND  ")])
    if includeInWhereClause(column_where, len(column_names) - 1):  # checks last column
        statement = ''.join(
            [statement, (str(column_names[len(column_names) - 1].value) + ' = ')])
        string = isValueTypeString(column_types, len(column_names) - 1)
        if string:  # add quotes
            statement = ''.join(
                [statement, ("'" + str(table[row][len(table[row]) - 1].value) + "';")])
        else:
            statement = ''.join(
                [statement, (str(table[row][len(table[row]) - 1].value) + ";")])
    else:
        # if last column is not in where clause drop 'AND' statement or drop ' WHERE' statement
        statement = statement[:-7] + ';'

    return statement


def createUpdateScripts(table_name, table):
    '''Creates the update scripts based on the data provided in the Excel spreadsheet.

    :param1 table_name: str
    :param2 table: tuple

    :return: dict
    '''

    script_dict = {}  # {cell: script}. ex. {'G7': 'UPDATE... ;'}
    column_names = table[cons.COLUMN_NAMES_ROW_INDEX]
    column_types = table[cons.COLUMN_DATA_TYPE_ROW_INDEX]
    column_includes = table[cons.INCLUDE_ROW_INDEX]
    column_where = table[cons.WHERE_ROW_INDEX]

    pre_statement = 'UPDATE ' + table_name + ' SET '

    # creates script for each row of data in the Excel table
    for row in range(cons.START_OF_DATA_ROWS_INDEX, len(table)):
        update_statement = createUpdateClause(
            table, column_names, column_types, column_includes, pre_statement, row)

        where_statement = createWhereClause(
            table, column_names, column_types, column_where, update_statement, row)

        excel_cell = excel_global.getExcelCellToInsertInto(
            len(table[row]), row)
        script_dict[excel_cell] = where_statement

    return script_dict


def createDeleteScripts(table_name, table):
    '''Creates the delete scripts based on the data provided in the Excel spreadsheet.

    :param1 table_name: str
    :param2 table: tuple

    :return: dict
    '''

    script_dict = {}  # {cell: script}. ex. {'G7': 'DELETE... ;'}
    column_names = table[cons.COLUMN_NAMES_ROW_INDEX]
    column_types = table[cons.COLUMN_DATA_TYPE_ROW_INDEX]
    column_where = table[cons.WHERE_ROW_INDEX]

    pre_statement = 'DELETE FROM ' + table_name + ' WHERE '

    # creates script for each row of data in the Excel table
    for row in range(cons.START_OF_DATA_ROWS_INDEX, len(table)):
        where_statement = createWhereClause(
            table, column_names, column_types, column_where, pre_statement, row)

        excel_cell = excel_global.getExcelCellToInsertInto(
            len(table[row]), row)
        script_dict[excel_cell] = where_statement

    return script_dict


def createSelectScripts(table_name, table):
    '''Creates the select scripts based on the data provided in the Excel spreadsheet.

    :param1 table_name: str
    :param2 table: tuple

    :return: dict
    '''

    script_dict = {}  # {cell: script}. ex. {'G7': 'INSERT INTO... ;'}
    column_names = table[cons.COLUMN_NAMES_ROW_INDEX]
    column_types = table[cons.COLUMN_DATA_TYPE_ROW_INDEX]
    column_includes = table[cons.INCLUDE_ROW_INDEX]
    column_where = table[cons.WHERE_ROW_INDEX]

    pre_statement = 'SELECT ('

    select_statement = createColumnClause(
        column_names, column_includes, pre_statement) + ') FROM ' + table_name + ' WHERE '

    # creates script for each row of data in the Excel table
    for row in range(cons.START_OF_DATA_ROWS_INDEX, len(table)):
        where_statement = createWhereClause(
            table, column_names, column_types, column_where, select_statement, row)

        excel_cell = excel_global.getExcelCellToInsertInto(
            len(table[row]), row)
        script_dict[excel_cell] = where_statement

    return script_dict


def saveToExcel(workbook):
    '''Saves the workbook to an Excel file.

    :param1 workbook: openpyxl.workbook.workbook.Workbook
    '''
    output_string = "Select/create the filename of Excel workbook you'd like to save/write to: "
    excel_global.createPopUpBox(
        output_string)  # tkinter dialog box

    file = tkinter.Tk()
    # opens file explorer so user can choose file to write to
    file.filename = tkFileDialog.asksaveasfilename(
        initialdir="C:/", title="Select/create file to save/write to", defaultextension=".xlsx")
    # saves new workbook with generated scripts to a user selected file
    workbook.save(file.filename)
    file.destroy()

    output_string = "Scripts saved to: '" + \
        str(file.filename) + "'"
    excel_global.createPopUpBox(
        output_string)  # tkinter dialog box


def writeToExcel(workbook, validate_with_sql):
    '''Iterates through each worksheet in the imported workbook, creates
    scripts for each worksheet, and writes the scripts to a new workbook. Returns
    True if scripts were generated and need to be saved, otherwise False

    :param1 workbook: openpyxl.workbook.workbook.Workbook
    :param2 validate_with_sql: str

    :return: bool
    '''

    any_changes = ''

    for worksheet in workbook.worksheets:
        if worksheet.title != 'configuration':  # skip the configuration sheet in the Excel book
            # check if worksheet is is valid and if user wants to write scripts for them
            valid_template = validate.validWorksheet(
                worksheet, validate_with_sql)

            if valid_template:  # only write to Excel if the Excel spreadsheet is a valid format
                script_type = getTypeOfScript(
                    worksheet)  # tkinter dialog box

                table_name = getTableName(worksheet)
                all_rows = tuple(worksheet.rows)

                # returns dict containing excel cell coordinates as key and script to writeas value
                scripts = writeScripts(
                    all_rows, script_type, table_name)

                any_changes = 'Excel'  # changes were made and need to be saved
                # writes script to worksheet
                for cell, script in scripts.items():
                    worksheet[cell] = script

    saveToExcel(workbook)

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
    excel_global.createPopUpBox(
        output_string)  # tkinter dialog box


def writeToSQL(workbook, validate_with_sql):
    '''Iterates through each worksheet in the imported workbook, creates
    scripts for each worksheet, and writes the scripts to a SQL file. Returns
    True if scripts were generated and have been saved, otherwise False

    :param1 workbook: openpyxl.workbook.workbook.Workbook
    :param2 validate_with_sql: str

    :return: bool
    '''

    any_changes = ''
    text_file = ''

    for worksheet in workbook.worksheets:
        if worksheet.title != 'configuration':  # skip the configuration sheet in the Excel book
            # check if worksheet is is valid and if user wants to write scripts for them
            valid_template = validate.validWorksheet(
                worksheet, validate_with_sql)

            if valid_template:  # only write to Excel if the Excel spreadsheet is a valid format
                script_type = getTypeOfScript(
                    worksheet)  # tkinter dialog box

                table_name = getTableName(worksheet)

                all_rows = tuple(worksheet.rows)

                # returns dict containing excel cell coordinates as key and script to writeas value
                scripts = writeScripts(
                    all_rows, script_type, table_name)

                any_changes = 'SQL'

                for cell, script in scripts.items():
                    text_file += script + '\n'

    saveToSQL(text_file)

    return any_changes
