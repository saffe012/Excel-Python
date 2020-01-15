'''
Module of 'excel.py' that handles functions related to generating an Excel
template for use in the script generation mode of this program
Matt Saffert
1-9-2020
'''

import constants as cons
import Tkinter
import excel_global


def populateIncludeRow(sql_table_name, column_names, column_is_nullable, column_is_identity, script_type):
    '''Creates Tkinter dialogue that asks user to check which data columns they'd
    like to include in their scripts. Run when script type is in (select, insert, update)

    :param1 sql_table_name: str
    :param2 column_names: List[str]
    :param3 column_is_nullable: List[str]
    :param4 column_is_identity: List[int]
    :param5 script_type: str

    :return: List[instance]
    '''

    include_values = []
    root = Tkinter.Tk()
    root.title('Excel Python')
    horizontal_sections = float(len(column_names) + 3)
    height = int(horizontal_sections * 50)
    wxh = "500x" + str(height)
    root.geometry(wxh)
    w = Tkinter.Label(
        root, text="Please select the columns you'd like to include in your script for the " + sql_table_name + " table:")
    w.pack()
    screen_fraction = 1 / horizontal_sections
    rely = screen_fraction
    w.place(relx=0.5, rely=rely, anchor='center')

    # for each column of data add a check box to dialog box to allow user to select or deselect
    for i in range(len(column_names)):
        rely = float('%.3f' % (rely + screen_fraction))
        var = Tkinter.IntVar()
        include_values.append(var)
        # not identity so column can be included in scripts
        if column_is_identity[i] == 0:
            # if script type is insert, and column cannot be null then automatically select
            if column_is_nullable[i] == 'NO' and script_type not in ('select', 'update'):
                b = Tkinter.Checkbutton(
                    root, text=column_names[i], variable=include_values[i], state='disabled')
                b.select()
                b.place(relx=0.5, rely=rely, anchor='center')
            else:  # if nullable or select or update, then data can be but does not need to be included
                b = Tkinter.Checkbutton(
                    root, text=column_names[i], variable=include_values[i])
                b.deselect()
                b.place(relx=0.5, rely=rely, anchor='center')
        else:  # column is identity column so cannot be updated or inserted into.
            if script_type != 'select':  # insert/update on identity column is NOT allowed
                b = Tkinter.Checkbutton(
                    root, text=column_names[i], variable=include_values[i], state='disabled')
                b.deselect()
                b.place(relx=0.5, rely=rely, anchor='center')
            else:  # select on identity column is allowed
                b = Tkinter.Checkbutton(
                    root, text=column_names[i], variable=include_values[i])
                b.deselect()
                b.place(relx=0.5, rely=rely, anchor='center')

    rely = float('%.3f' % (rely + screen_fraction))
    button = Tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=rely, anchor='center')
    Tkinter.mainloop()

    return include_values


def populateWhereRow(sql_table_name, column_names, script_type):
    '''Creates Tkinter dialogue that asks user to check which columns they'd
    like to include in the WHERE clause of their scripts where script_type
    in (update, select, delete)

    :param1 sql_table_name: str
    :param2 column_names: List[str]
    :param3 script_type: str

    :return: List[instance]
    '''

    where_values = []
    root = Tkinter.Tk()
    root.title('Excel Python')
    horizontal_sections = float(len(column_names) + 3)
    height = int(horizontal_sections * 50)
    wxh = "650x" + str(height)
    root.geometry(wxh)
    w = Tkinter.Label(
        root, text="Please select the columns you'd like have in the where clause of your script for the " + sql_table_name + " table:")
    w.pack()
    screen_fraction = 1 / horizontal_sections
    rely = screen_fraction
    w.place(relx=0.5, rely=rely, anchor='center')

    # for each column of data add a check box to dialog box to allow user to select or deselect
    for i in range(len(column_names)):
        rely = float('%.3f' % (rely + screen_fraction))
        var = Tkinter.IntVar()
        where_values.append(var)
        b = Tkinter.Checkbutton(
            root, text=column_names[i], variable=where_values[i])
        b.deselect()
        b.place(relx=0.5, rely=rely, anchor='center')

    rely = float('%.3f' % (rely + screen_fraction))
    button = Tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=rely, anchor='center')

    Tkinter.mainloop()

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

    :return: List[instance], List[instance]
    '''

    if script_type in ('select', 'insert', 'update'):  # generate include row
        sql_include_row = populateIncludeRow(
            sql_table_name, sql_column_names, column_is_nullable, column_is_identity, script_type)
    else:
        sql_include_row = []

    if script_type in ('select', 'delete', 'update'):  # generate where row
        sql_where_row = populateWhereRow(
            sql_table_name, sql_column_names, script_type)
    else:
        sql_where_row = []

    return sql_include_row, sql_where_row


def WriteTemplateToSheet(worksheet, sql_table_name, script_type, sql_column_names, sql_column_types, sql_include_row, sql_where_row):
    '''Based on user input from previous functions, this function will write the
    template to the Excel workbook.

    :param1 worksheet: openpyxl.worksheet.worksheet.Worksheet
    :param2 sql_table_name: str
    :param3 script_type: str
    :param4 sql_column_names: List[str]
    :param5 sql_column_types: List[str]
    :param6 sql_include_row: List[instance]
    :param7 sql_where_row: List[instance]

    :return: NONE
    '''

    # populates top info row
    worksheet[excel_global.getExcelCellToInsertInto(
        0, cons.INFO_ROW)] = sql_table_name
    worksheet[excel_global.getExcelCellToInsertInto(1, cons.INFO_ROW)] = script_type

    # populates next 4 rows in the Excel template with data from column lists
    for i in range(len(sql_column_names)):
        worksheet[excel_global.getExcelCellToInsertInto(
            i, cons.COLUMN_NAMES_ROW_INDEX)] = sql_column_names[i]
        worksheet[excel_global.getExcelCellToInsertInto(
            i, cons.COLUMN_DATA_TYPE_ROW_INDEX)] = sql_column_types[i]
        if len(sql_include_row) > 0:  # only put data in include row if there id data
            if sql_include_row[i].get() == 1:
                worksheet[excel_global.getExcelCellToInsertInto(
                    i, cons.INCLUDE_ROW_INDEX)] = 'include'
        if len(sql_where_row) > 0:  # only put data in where row if there id data
            if sql_where_row[i].get() == 1:
                worksheet[excel_global.getExcelCellToInsertInto(
                    i, cons.WHERE_ROW_INDEX)] = 'where'


def getTypeOfScriptFromUser(worksheet_title):
    '''Creates a Tkinter dialog box that asks the user to choose the type of scripts
    they are trying to generate

    :return: instance
    '''

    root = Tkinter.Tk()
    root.title('Excel Python')
    root.geometry("500x500")
    script_type = Tkinter.StringVar()
    script_type.set("insert")
    w = Tkinter.Label(
        root, text="Please choose what type of scripts you'd like to create for '" + worksheet_title + "' worksheet:")
    w.pack()
    w.place(relx=0.5, rely=0.1, anchor='center')
    Tkinter.Radiobutton(root, text='insert', variable=script_type,
                        value='insert').place(relx=0.5, rely=0.2, anchor='center')
    Tkinter.Radiobutton(root, text='update', variable=script_type,
                        value='update').place(relx=0.5, rely=0.3, anchor='center')
    Tkinter.Radiobutton(root, text='delete', variable=script_type,
                        value='delete').place(relx=0.5, rely=0.4, anchor='center')
    Tkinter.Radiobutton(root, text='select', variable=script_type,
                        value='select').place(relx=0.5, rely=0.5, anchor='center')
    button = Tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.6, anchor='center')
    Tkinter.mainloop()

    return script_type
