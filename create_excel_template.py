'''
Module of 'excel.py' that handles functions related to generating an Excel
template for use in the script generation mode of this program
Matt Saffert
1-9-2020
'''

import constants as cons
import tkinter
import excel_global
import subprocess


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
    root = tkinter.Tk()
    excel_global.addQuitMenuButton(root)
    root.title('Excel Python')
    if len(column_names) < 10:
        horizontal_sections = float(len(column_names) + 3)
        height = int(horizontal_sections * 50)
    else:
        horizontal_sections = 13.0
        height = 550
    width = 500 + ((len(column_names) // 11) * 150)
    vertical_sections = float((len(column_names) // 11) + 2)
    wxh = str(width) + "x" + str(height)
    root.geometry(wxh)
    w = tkinter.Label(
        root, text="Please select the columns you'd like to include in your script for the " + sql_table_name + " table:")
    w.pack()
    vertical_screen_fraction = 1 / vertical_sections
    relx = float('%.3f' % (vertical_screen_fraction))
    horizontal_screen_fraction = 1 / horizontal_sections
    rely = horizontal_screen_fraction
    w.place(relx=0.5, rely=rely, anchor='center')
    count = 0

    # for each column of data add a check box to dialog box to allow user to select or deselect
    for i in range(len(column_names)):
        if count >= 10:
            count = 0
            relx = float('%.3f' % (relx + vertical_screen_fraction))
            rely = horizontal_screen_fraction
        rely = float('%.3f' % (rely + horizontal_screen_fraction))
        var = tkinter.IntVar()
        disable_change.append(0)
        # not identity so column can be included in scripts
        if column_is_identity[i] == 0:
            # if script type is insert, and column cannot be null then automatically select
            if column_is_nullable[i] == 'NO' and script_type not in ('select', 'update'):
                include_values.append(var)
                b = tkinter.Checkbutton(
                    root, text=column_names[i], variable=include_values[i], state='disabled')
                disable_change[i] = 1
                b.select()
                b.place(relx=relx, rely=rely, anchor='center')
            else:  # if nullable or select or update, then data can be but does not need to be included
                include_values.append(var)
                b = tkinter.Checkbutton(
                    root, text=column_names[i], variable=include_values[i])
                b.deselect()
                b.place(relx=relx, rely=rely, anchor='center')
        else:  # column is identity column so cannot be updated or inserted into.
            if script_type != 'select':  # insert/update on identity column is NOT allowed
                include_values.append(var)
                b = tkinter.Checkbutton(
                    root, text=column_names[i], variable=include_values[i], state='disabled')
                disable_change[i] = 1
                b.deselect()
                b.place(relx=relx, rely=rely, anchor='center')
            else:  # select on identity column is allowed
                include_values.append(var)
                b = tkinter.Checkbutton(
                    root, text=column_names[i], variable=include_values[i])
                b.deselect()
                b.place(relx=relx, rely=rely, anchor='center')
        count += 1

    rely = float('%.3f' % (rely + horizontal_screen_fraction))
    if len(column_names) < 10:
        button = tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
            relx=0.5, rely=rely, anchor='center')
    else:
        button = tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
            relx=0.5, rely=(horizontal_screen_fraction * 12), anchor='center')

    tkinter.mainloop()

    for i in range(len(include_values)):
        include_values[i] = include_values[i].get()

    return include_values, disable_change


def populateWhereRow(sql_table_name, column_names, script_type):
    '''Creates tkinter dialogue that asks user to check which columns they'd
    like to include in the WHERE clause of their scripts where script_type
    in (update, select, delete)

    :param1 sql_table_name: str
    :param2 column_names: List[str]
    :param3 script_type: str

    :return: List[int]
    '''

    where_values = []
    root = tkinter.Tk()
    excel_global.addQuitMenuButton(root)
    root.title('Excel Python')

    if len(column_names) < 10:
        horizontal_sections = float(len(column_names) + 3)
        height = int(horizontal_sections * 50)
    else:
        horizontal_sections = 13.0
        height = 550
    width = 500 + ((len(column_names) // 11) * 150)
    vertical_sections = float((len(column_names) // 11) + 2)
    wxh = str(width) + "x" + str(height)

    root.geometry(wxh)
    w = tkinter.Label(
        root, text="Please select the columns you'd like have in the where clause of your script for the " + sql_table_name + " table:")
    w.pack()

    vertical_screen_fraction = 1 / vertical_sections
    relx = float('%.3f' % (vertical_screen_fraction))
    horizontal_screen_fraction = 1 / horizontal_sections
    rely = horizontal_screen_fraction
    w.place(relx=0.5, rely=rely, anchor='center')
    count = 0

    # for each column of data add a check box to dialog box to allow user to select or deselect
    for i in range(len(column_names)):
        if count >= 10:
            count = 0
            relx = float('%.3f' % (relx + vertical_screen_fraction))
            rely = horizontal_screen_fraction
        rely = float('%.3f' % (rely + horizontal_screen_fraction))
        var = tkinter.IntVar()
        where_values.append(var)
        b = tkinter.Checkbutton(
            root, text=column_names[i], variable=where_values[i])
        b.deselect()
        b.place(relx=relx, rely=rely, anchor='center')
        count += 1

    rely = float('%.3f' % (rely + horizontal_screen_fraction))
    if len(column_names) < 10:
        button = tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
            relx=0.5, rely=rely, anchor='center')
    else:
        button = tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
            relx=0.5, rely=(horizontal_screen_fraction * 12), anchor='center')

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
            sql_table_name, sql_column_names, script_type)
    else:
        sql_where_row = []

    return sql_include_row, sql_where_row, disable_include_change


def WriteTemplateToSheet(worksheet, sql_table_name, script_type, sql_column_names, sql_column_types, sql_include_row, sql_where_row, disable_include_change):
    '''Based on user input from previous functions, this function will write the
    template to the Excel workbook.

    :param1 worksheet: openpyxl.worksheet.worksheet.Worksheet
    :param2 sql_table_name: str
    :param3 script_type: str
    :param4 sql_column_names: List[str]
    :param5 sql_column_types: List[str]
    :param6 sql_include_row: List[int]
    :param7 sql_where_row: List[int]
    :param8 disable_include_change: List[int]

    :return: NONE
    '''

    # populates top info row
    worksheet[excel_global.getExcelCellToInsertInto(
        0, cons.INFO_ROW)] = sql_table_name
    worksheet[excel_global.getExcelCellToInsertInto(
        1, cons.INFO_ROW)] = script_type

    # populates next 4 rows in the Excel template with data from column lists
    for i in range(len(sql_column_names)):
        worksheet[excel_global.getExcelCellToInsertInto(
            i, cons.COLUMN_NAMES_ROW_INDEX)] = sql_column_names[i]
        worksheet[excel_global.getExcelCellToInsertInto(
            i, cons.COLUMN_DATA_TYPE_ROW_INDEX)] = sql_column_types[i]
        if len(sql_include_row) > 0:  # only put data in include row if there is data
            cell = excel_global.getExcelCellToInsertInto(
                i, cons.INCLUDE_ROW_INDEX)
            if sql_include_row[i] == 1:
                worksheet[cell] = 'include'
            # if the cell shouldn't be changed, color it red
            if disable_include_change[i] == 1:
                excel_global.colorCell(worksheet, cell, cons.RED)
        else:  # if script is delete, there should be no include. color it red
            excel_global.colorCell(worksheet, excel_global.getExcelCellToInsertInto(
                i, cons.INCLUDE_ROW_INDEX), cons.RED)
        if len(sql_where_row) > 0:  # only put data in where row if there is data
            if sql_where_row[i] == 1:
                worksheet[excel_global.getExcelCellToInsertInto(
                    i, cons.WHERE_ROW_INDEX)] = 'where'
        else:  # if script is insert, there should be no where clause. color it red
            excel_global.colorCell(worksheet, excel_global.getExcelCellToInsertInto(
                i, cons.WHERE_ROW_INDEX), cons.RED)


def getTypeOfScriptFromUser(worksheet_title):
    '''Creates a tkinter dialog box that asks the user to choose the type of scripts
    they are trying to generate

    :return: instance
    '''

    root = tkinter.Tk()
    excel_global.addQuitMenuButton(root)
    root.title('Excel Python')
    root.geometry("500x500")
    script_type = tkinter.StringVar()
    script_type.set("insert")
    w = tkinter.Label(
        root, text="Please choose what type of scripts you'd like to create for '" + worksheet_title + "' worksheet:")
    w.pack()
    w.place(relx=0.5, rely=0.1, anchor='center')
    tkinter.Radiobutton(root, text='insert', variable=script_type,
                        value='insert').place(relx=0.5, rely=0.2, anchor='center')
    tkinter.Radiobutton(root, text='update', variable=script_type,
                        value='update').place(relx=0.5, rely=0.3, anchor='center')
    tkinter.Radiobutton(root, text='delete', variable=script_type,
                        value='delete').place(relx=0.5, rely=0.4, anchor='center')
    tkinter.Radiobutton(root, text='select', variable=script_type,
                        value='select').place(relx=0.5, rely=0.5, anchor='center')
    button = tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.6, anchor='center')
    tkinter.mainloop()

    return script_type


def getTemplateType():
    '''Creates a tkinter dialog box that asks the user to choose the type of template
    they are trying to generate

    :return: str
    '''

    root = tkinter.Tk()
    excel_global.addQuitMenuButton(root)
    root.title('Excel Python')
    root.geometry("500x500")
    template_type = tkinter.StringVar()
    template_type.set("generic")
    w = tkinter.Label(
        root, text="Please choose what type of template you'd like to create:")
    w.pack()
    w.place(relx=0.5, rely=0.1, anchor='center')
    w = tkinter.Label(
        root, text="Generic templates should be edited in order to match the data you put into the template.")
    w.pack()
    w.place(relx=0.5, rely=0.2, anchor='center')
    tkinter.Radiobutton(root, text='Generic template', variable=template_type,
                        value='generic').place(relx=0.5, rely=0.3, anchor='center')
    tkinter.Radiobutton(root, text='Template from existing SQL table', variable=template_type,
                        value='from_table').place(relx=0.5, rely=0.4, anchor='center')
    button = tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.6, anchor='center')
    tkinter.mainloop()

    return template_type.get()


def generateGenericTemplate(worksheet):
    '''Based on user input from previous functions, this function will write the
    template to the Excel workbook.

    :param1 worksheet: openpyxl.worksheet.worksheet.Worksheet

    :return: str, str, List[str], List[str], List[int], List[int], List[int]
    '''

    sql_table_name = "IOChannels"
    script_type = 'insert'
    sql_column_names = ['Id', 'IOServersId', 'Name', 'HealthStatusAddressId',
                        'IsHealthStatusAddress', 'SimulateIO', 'IsConnected']
    sql_column_types = ['int', 'int',
                        'varchar(50)', 'int', 'bit', 'bit', 'bit']
    sql_include_row = [1, 1, 1, 1, 1, 1, 1]
    sql_where_row = [0, 0, 1, 0, 0, 0, 1]
    disable_include_change = [1, 0, 0, 0, 0, 0, 0]

    return sql_table_name, script_type, sql_column_names, sql_column_types, sql_include_row, sql_where_row, disable_include_change


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
    sql_table_name = excel_global.createDropDownBox(
        description, label, sql_tables)

    sql_column_names, sql_column_types, column_is_nullable, column_is_identity = excel_global.getSQLTableInfo(sql_table_name, cursor)

    return sql_column_names, sql_column_types, column_is_nullable, column_is_identity, sql_table_name
