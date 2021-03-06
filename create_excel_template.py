'''
Module of 'excel.py' that handles functions related to generating an Excel
template for use in the script generation mode of this program
Matt Saffert
1-9-2020
'''

import excel_global
import tkinter
import pandas as pd
from excel_constants import *
import global_gui as gui


def templateMode():
    '''Starts the template generation mode of the application.

    :return: NONE
    '''

    template_type = getTemplateType()

    if template_type == 'from_table':  # generates an Excel template from a SQL database
        workbook = templateModeSQL()

    elif template_type == 'generic':  # generates a generic template with default table data
        # dictionary filled with generic data to build template
        workbook = {'IOChannels': pd.DataFrame(data=GENERIC_TEMPLATE)}

    else:
        gui.closeProgram()

    gui.saveToExcel(workbook)


def populateIncludeRow(sql_table_name, column_names, column_is_nullable, column_is_identity, script_type):
    '''Creates tkinter dialogue that asks user to check which data columns they'd
    like to include in their scripts. Run when script type is in (select, insert, update)

    :param1 sql_table_name: str
    :param2 column_names: List[str]
    :param3 column_is_nullable: List[str]
    :param4 column_is_identity: List[int]
    :param5 script_type: str

    :return: List[int], List[int]
    '''

    include_values = []
    disable_change = []

    width_x_height, x_spacing, y_spacing, vertical_screen_fraction, horizontal_screen_fraction = calculateGUISpacing(
        column_names)

    description = "Please select the columns you'd like to include in your script for the " + \
        sql_table_name + " table:"
    root = gui.generateWindow(
        width_x_height, description, relx=0.5, rely=y_spacing)

    count = 0

    # for each column of data add a check box to dialog box to allow user to select or deselect
    for i in range(len(column_names)):
        if count >= 10:
            count = 0
            x_spacing = float('%.3f' % (x_spacing + vertical_screen_fraction))
            y_spacing = horizontal_screen_fraction
        y_spacing = float('%.3f' % (y_spacing + horizontal_screen_fraction))
        var = tkinter.IntVar()
        disable_change.append(0)
        # not identity so column can be included in scripts
        if column_is_identity[i] == 0:
            # if script type is insert, and column cannot be null then automatically select
            if column_is_nullable[i] == 'NO' and script_type not in ('select', 'update'):
                include_values.append(var)
                disable_change[i] = 1
                gui.createCheckBox(root, column_names[i], include_values[i], x_spacing, y_spacing, select=True, disable='disabled')
            else:  # if nullable or select or update, then data can be but does not need to be included
                include_values.append(var)
                gui.createCheckBox(root, column_names[i], include_values[i], x_spacing, y_spacing, select=False, disable='normal')
        else:  # column is identity column so cannot be updated or inserted into.
            if script_type != 'select':  # insert/update on identity column is NOT allowed
                include_values.append(var)
                disable_change[i] = 1
                gui.createCheckBox(root, column_names[i], include_values[i], x_spacing, y_spacing, select=False, disable='disabled')
            else:  # select on identity column is allowed
                include_values.append(var)
                gui.createCheckBox(root, column_names[i], include_values[i], x_spacing, y_spacing, select=False, disable='normal')
        count += 1

    placeNextButton(y_spacing, horizontal_screen_fraction, root, column_names)

    tkinter.mainloop()

    for i in range(len(include_values)):
        include_values[i] = include_values[i].get()

    return include_values, disable_change


def populateWhereRow(sql_table_name, column_names):
    '''Creates tkinter dialogue that asks user to check which columns they'd
    like to include in the WHERE clause of their scripts where script_type
    in (update, select, delete)

    :param1 sql_table_name: str
    :param2 column_names: List[str]

    :return: List[int]
    '''

    where_values = []

    width_x_height, x_spacing, y_spacing, vertical_screen_fraction, horizontal_screen_fraction = calculateGUISpacing(
        column_names)

    description = "Please select the columns you'd like have in the where clause of your script for the " + \
        sql_table_name + " table:"
    root = gui.generateWindow(
        width_x_height, description, relx=0.5, rely=y_spacing)

    count = 0

    # for each column of data add a check box to dialog box to allow user to select or deselect
    for i in range(len(column_names)):
        if count >= 10:
            count = 0
            x_spacing = float('%.3f' % (x_spacing + vertical_screen_fraction))
            y_spacing = horizontal_screen_fraction
        y_spacing = float('%.3f' % (y_spacing + horizontal_screen_fraction))
        var = tkinter.IntVar()
        where_values.append(var)

        gui.createCheckBox(root, column_names[i], where_values[i], x_spacing, y_spacing, select=False, disable='normal')
        count += 1

    placeNextButton(y_spacing, horizontal_screen_fraction, root, column_names)

    tkinter.mainloop()

    for i in range(len(where_values)):
        where_values[i] = where_values[i].get()

    return where_values


def populateClauses(sql_table_name, sql_column_names, column_is_nullable, column_is_identity, script_type):
    '''Depending on the type of scripts that are being generated, this function
    calls the appropriate routines to populate the include and/or where row
    of the Excel spreadsheet

    :param1 sql_table_name: str
    :param2 sql_column_names: List[str]
    :param3 column_is_nullable: List[str]
    :param4 column_is_identity: List[int]
    :param5 script_type: str

    :return: List[int], List[int], List[int]
    '''

    if script_type in ('select', 'insert', 'update'):  # generate include row
        sql_include_row, disable_include_change = populateIncludeRow(
            sql_table_name, sql_column_names, column_is_nullable, column_is_identity, script_type)
    else:
        disable_include_change = []
        sql_include_row = []

    if script_type in ('select', 'delete', 'update'):  # generate where row
        sql_where_row = populateWhereRow(
            sql_table_name, sql_column_names)
    else:
        sql_where_row = []

    return sql_include_row, sql_where_row, disable_include_change


def WriteTemplateToSheet(worksheet, sql_column_names, sql_column_types, sql_include_row, sql_where_row, disable_include_change):
    '''Based on user input from previous functions, this function will write the
    template to the Excel workbook.

    :param1 worksheet: pandas.core.frame.DataFrame
    :param2 sql_column_names: List[str]
    :param3 sql_column_types: List[str]
    :param4 sql_include_row: List[int]
    :param5 sql_where_row: List[int]
    :param6 disable_include_change: List[int]

    :return: NONE
    '''

    # populates top info row
    worksheet.iloc[INFO_ROW][TABLE_NAME] = sql_table_name
    worksheet.iloc[INFO_ROW][SCRIPT_TYPE] = script_type

    # populates next 4 rows in the Excel template with data from column lists
    for i in range(len(sql_column_names)):
        worksheet.iloc[COLUMN_NAMES_ROW_INDEX][i] = sql_column_names[i]
        worksheet.iloc[COLUMN_DATA_TYPE_ROW_INDEX][i] = sql_column_types[i]
        if len(sql_include_row) > 0:  # only put data in include row if there is data
            if sql_include_row[i] == 1:
                worksheet.iloc[INCLUDE_ROW_INDEX][i] = 'include'
            # if the cell shouldn't be changed, color it red
            if disable_include_change[i] == 1:
                worksheet.iloc[INCLUDE_ROW_INDEX][i] = worksheet.iloc[INCLUDE_ROW_INDEX][i].upper(
                )
        else:  # if script is delete, there should be no include. color it red
            worksheet.iloc[INCLUDE_ROW_INDEX][i] = worksheet.iloc[INCLUDE_ROW_INDEX][i].upper(
            )
        if len(sql_where_row) > 0:  # only put data in where row if there is data
            if sql_where_row[i] == 1:
                worksheet.iloc[WHERE_ROW_INDEX][i] = 'where'
        else:  # if script is insert, there should be no where clause. color it red
            worksheet.iloc[INCLUDE_ROW_INDEX][i] = worksheet.iloc[INCLUDE_ROW_INDEX][i].upper(
            )


def getTypeOfScriptFromUser(worksheet_title):
    '''Creates a tkinter dialog box that asks the user to choose the type of scripts
    they are trying to generate

    :param1 worksheet_title: str

    :return: instance
    '''

    description = "Please choose what type of scripts you'd like to create for '" + \
        worksheet_title + "' worksheet:"
    root = gui.generateWindow("500x500", description, relx=0.5, rely=0.1)

    script_type = tkinter.StringVar()
    script_type.set("insert")

    tkinter.Radiobutton(root, text='insert', variable=script_type,
                        value='insert').place(relx=0.5, rely=0.2, anchor='center')
    tkinter.Radiobutton(root, text='update', variable=script_type,
                        value='update').place(relx=0.5, rely=0.3, anchor='center')
    tkinter.Radiobutton(root, text='delete', variable=script_type,
                        value='delete').place(relx=0.5, rely=0.4, anchor='center')
    tkinter.Radiobutton(root, text='select', variable=script_type,
                        value='select').place(relx=0.5, rely=0.5, anchor='center')
    tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.6, anchor='center')
    tkinter.mainloop()

    return script_type


def getTemplateType():
    '''Creates a tkinter dialog box that asks the user to choose the type of template
    they are trying to generate

    :return: str
    '''

    # Create box and add label
    description1 = "Please choose what type of template you'd like to create:"
    root = gui.generateWindow("500x500", description1, relx=0.5, rely=0.1)

    # Add second label
    description2 = "Generic templates should be edited in order to match the data you put into the template."
    gui.addLabelToBox(root, description2, 0.5, 0.2)

    template_type = tkinter.StringVar()
    template_type.set("generic")

    tkinter.Radiobutton(root, text='Generic template', variable=template_type,
                        value='generic').place(relx=0.5, rely=0.3, anchor='center')
    tkinter.Radiobutton(root, text='Template from existing SQL table', variable=template_type,
                        value='from_table').place(relx=0.5, rely=0.4, anchor='center')
    tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.6, anchor='center')
    tkinter.mainloop()

    return template_type.get()


def getTemplateInfo():
    '''Creates a series of tkinter dialogues that asks user to info about the
    template they are trying to create.

    :return: List[str], List[str], List[str], List[int], str
    '''

    # gets the name of the SQL instance from user, SQL DB from user, and list of tables in that DB
    sql_tables, cursor, sql_database_name = excel_global.connectToSQLServer()

    description = "Please enter the name of the table you'd like to work with in the " + \
        sql_database_name + " database:"
    label = 'SQL table name: '
    sql_table_name = gui.createDropDownBox(
        description, label, sql_tables)

    sql_column_names, sql_column_types, column_is_nullable, column_is_identity = excel_global.getSQLTableInfo(
        sql_table_name, cursor)

    return sql_column_names, sql_column_types, column_is_nullable, column_is_identity, sql_table_name


def calculateGUISpacing(column_names):
    '''Calculates the spacing required to place elements on the GUI box.

    :param1 column_names: List[str]

    :return: str, float, float, float, float
    '''

    if len(column_names) < 10:
        horizontal_sections = float(len(column_names) + 3)
        height = int(horizontal_sections * 50)
    else:
        horizontal_sections = 13.0
        height = 550

    width = 500 + ((len(column_names) // 11) * 150)
    vertical_sections = float((len(column_names) // 11) + 2)
    width_x_height = str(width) + "x" + str(height)

    vertical_screen_fraction = 1 / vertical_sections
    x_spacing = float('%.3f' % (vertical_screen_fraction))
    horizontal_screen_fraction = 1 / horizontal_sections
    y_spacing = horizontal_screen_fraction

    return width_x_height, x_spacing, y_spacing, vertical_screen_fraction, horizontal_screen_fraction


def templateModeSQL():
    '''Runs the template generation mode using SQL.

    :return: dict
    '''

    # tkinter dialog boxes
    sql_column_names, sql_column_types, column_is_nullable, column_is_identity, sql_table_name = getTemplateInfo()
    workbook = {sql_table_name: pd.DataFrame()}
    worksheet = workbook[sql_table_name]

    # allows user to select the type of script this template is for
    script_type = getTypeOfScriptFromUser(
        sql_table_name).get()  # tkinter dialog box

    # asks user which elements from the imported table they'd like to include in their scripts
    sql_include_row, sql_where_row, disable_include_change = populateClauses(
        sql_table_name, sql_column_names, column_is_nullable, column_is_identity, script_type)  # tkinter dialog boxes

    # writes the generated template to the new Excel workbook
    WriteTemplateToSheet(
        worksheet, sql_column_names, sql_column_types, sql_include_row, sql_where_row, disable_include_change)

    return workbook


def placeNextButton(y_spacing, horizontal_screen_fraction, root, column_names):
    '''Places "Next" button on GUI that will advance program.

    :param1 y_spacing: float
    :param2 horizontal_screen_fraction: float
    :param3 root: tkinter
    :param4 column_names: List[str]

    :return: NONE
    '''

    y_spacing = float('%.3f' % (y_spacing + horizontal_screen_fraction))
    if len(column_names) < 10:
        tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
            relx=0.5, rely=y_spacing, anchor='center')
    else:
        tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
            relx=0.5, rely=(horizontal_screen_fraction * 12), anchor='center')
