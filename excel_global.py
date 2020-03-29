'''
Module of 'excel.py' that contains global functions.
Matt Saffert
1-9-2020
'''

import tkinter
import pyodbc
import constants as cons
import subprocess
import sys
from tkinter import filedialog as tkFileDialog
import pandas as pd


def openExcelFile(output_string):
    '''Opens an existing Excel workbook using pandas.

    :param1 output_string: str

    :return: dict
    '''

    createPopUpBox(output_string)  # tkinter dialog box

    file = tkinter.Tk()
    # opens file explorer so user can choose file to read from
    file.filename = tkFileDialog.askopenfilename(
        initialdir="C:/", title="Select file to write scripts for")
    file.destroy()
    workbook = pd.read_excel(file.filename, header=None, sheet_name=None)
    for worksheet in workbook:
        workbook[worksheet] = workbook[worksheet].rename(index={0: "info"})
        workbook[worksheet] = workbook[worksheet].rename(index={1: "names"})
        workbook[worksheet] = workbook[worksheet].rename(index={2: "types"})
        workbook[worksheet] = workbook[worksheet].rename(index={3: "include"})
        workbook[worksheet] = workbook[worksheet].rename(index={4: "where"})
        for i in range(5, len(workbook[worksheet])):
            workbook[worksheet] = workbook[worksheet].rename(index={
                                                             i: (i - 5)})

    return workbook


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
        sql_server_name = createTextEntryBox(
            description, label).get()
        if sql_server_name == '':
            createPopUpBox(
                "Please enter a SQL server instance name.")
    '''
    '''
    description = "Please choose the name of the SQL Server where your database is located:"
    label = 'SQL Server name: '
    sql_server_name = createDropDownBox(description, label, local_servers)
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

    sql_database_name = createDropDownBox(
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


def addQuitMenuButton(root):
    '''Adds quiting capability to a tkinter box both as menu option and the "X"
    in upper right hand corner of box

    :param1 root: tkinter

    :return: NONE
    '''

    menubar = tkinter.Menu(root)
    menubar.add_command(label="Quit!", command=lambda: closeProgram())
    root.protocol("WM_DELETE_WINDOW", lambda: closeProgram())
    root.config(menu=menubar)


def closeProgram():
    '''Closes the program after confirming with user

    :return: NONE
    '''

    root = tkinter.Tk()
    addQuitMenuButton(root)
    root.title('Excel Python')
    output_string = "Are you sure you want to close the program?"
    w = tkinter.Label(root, text=output_string)
    w.pack()
    w.place(relx=0.5, rely=0.2, anchor='center')
    str_len = len(output_string)
    text_height = (str_len // 35) + 1
    height = 150 + (text_height * 10)
    root.geometry("450x150")
    button = tkinter.Button(root, text='Yes', width=15, command=lambda: quit()).place(
        relx=0.35, rely=0.8, anchor='center')
    button = tkinter.Button(root, text='No', width=15, command=root.destroy).place(
        relx=0.65, rely=0.8, anchor='center')
    tkinter.mainloop()


def createYesNoBox(description, label1, label2):
    '''Creates a tkinter pop-up box that gives the user a choice between 2 options

    :param1 description: str
    :param2 label1: str
    :param3 label2: str

    :return: str
    '''

    root = tkinter.Tk()
    addQuitMenuButton(root)
    root.title('Excel Python')
    root.geometry("500x500")
    program_mode = tkinter.StringVar()
    program_mode.set(label1)
    w = tkinter.Label(
        root, text=description)
    w.pack()
    w.place(relx=0.5, rely=0.1, anchor='center')
    tkinter.Radiobutton(root, text=label1, variable=program_mode,
                        value=label1).place(relx=0.5, rely=0.4, anchor='center')
    tkinter.Radiobutton(root, text=label2, variable=program_mode,
                        value=label2).place(relx=0.5, rely=0.5, anchor='center')
    button = tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.6, anchor='center')
    tkinter.mainloop()

    return program_mode.get()


def getExcelCellToInsertInto(column, row):
    '''Gets the column and row of the excel spreadsheet that the script should be inserted into.

    :param1 column: int
    :param2 row: int

    :return: str
    '''

    # column is retrieved by finding the key of the LETTER_INDEX_DICT that the value(index) belongs to.
    excel_column = list(cons.LETTER_INDEX_DICT.keys())[list(
        cons.LETTER_INDEX_DICT.values()).index(column)]
    excel_row = str(row + 1)
    # excel coordinate cell that script should be inserted into
    excel_cell = excel_column + excel_row

    return excel_cell


def getProgramMode():
    '''Creates a tkinter dialog box that asks the user what mode they'd like
    the program to enter. (build an Excel template or write scripts)

    :return: instance
    '''

    root = tkinter.Tk()
    addQuitMenuButton(root)
    root.title('Excel Python')
    root.geometry("500x500")
    program_mode = tkinter.StringVar()
    program_mode.set("scripts")
    w = tkinter.Label(
        root, text="Would you like to build an Excel template or write SQL scripts to an Excel file: ")
    w.pack()
    w.place(relx=0.5, rely=0.1, anchor='center')
    tkinter.Radiobutton(root, text='Check if workbook is valid for writing scripts', variable=program_mode,
                        value='validate').place(relx=0.5, rely=0.3, anchor='center')
    tkinter.Radiobutton(root, text='Build Excel template', variable=program_mode,
                        value='template').place(relx=0.5, rely=0.4, anchor='center')
    tkinter.Radiobutton(root, text='Write SQL scripts', variable=program_mode,
                        value='scripts').place(relx=0.5, rely=0.5, anchor='center')
    button = tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.6, anchor='center')
    tkinter.mainloop()

    return program_mode


def createPopUpBox(output_string):
    '''Creates a tkinter pop-up box that displays whatever test is input with an "Ok" button
    to acknowledge info/close window

    :param1 output_string: str
    '''

    root = tkinter.Tk()
    addQuitMenuButton(root)
    root.title('Excel Python')
    w = tkinter.Label(root, text=output_string)
    w.pack()
    w.place(relx=0.5, rely=0.2, anchor='center')
    str_len = len(output_string)
    text_height = (str_len // 35) + 1
    height = 150 + (text_height * 10)
    root.geometry("450x150")
    button = tkinter.Button(root, text='Ok', width=25, command=root.destroy).place(
        relx=0.5, rely=0.8, anchor='center')
    tkinter.mainloop()


def createErrorBox(output_string):
    '''Creates a tkinter pop-up box that displays the error message with an "Ok" button
    to acknowledge error/close window

    :param1 output_string: str
    '''

    root = tkinter.Tk()
    addQuitMenuButton(root)
    root.title('Excel Python')
    str_len = len(output_string)
    text_height = (str_len // 35) + 1
    height = 150 + (text_height * 10)
    root.geometry("450x150")
    T = tkinter.Text(root, height=text_height, width=35)
    T.pack()
    T.insert(tkinter.END, output_string)
    button = tkinter.Button(root, text='Ok', width=25, command=root.destroy).place(
        relx=0.5, rely=0.8, anchor='center')
    tkinter.mainloop()


def createTextEntryBox(description, label):
    '''Creates a tkinter dialog box that asks the user to enter the requested
    information in a text box.

    :param1 description: str
    :param2 label: str

    :return: instance
    '''

    root = tkinter.Tk()
    addQuitMenuButton(root)
    root.title('Excel Python')
    root.geometry("600x400")
    entry_value = tkinter.StringVar()
    w = tkinter.Label(root, text=description)
    w.pack()
    w.place(relx=0.5, rely=0.3, anchor='center')
    tkinter.Label(root, text=label).place(
        relx=0.4, rely=0.4, anchor='center')
    e1 = tkinter.Entry(root, textvariable=entry_value)
    e1.place(relx=0.6, rely=0.4, anchor='center')
    button = tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.5, anchor='center')
    tkinter.mainloop()

    return entry_value


def createDropDownBox(description, label, data):
    '''Creates a tkinter dialog box that asks the user to enter the requested
    information in a text box.

    :param1 description: str
    :param2 label: str
    :param3 data: List[?]

    :return: ?
    '''

    root = tkinter.Tk()
    addQuitMenuButton(root)
    root.title('Excel Python')
    root.geometry("500x500")

    value = tkinter.StringVar(root)
    value.set(data[0])  # default value

    w = tkinter.Label(root, text=description)
    w.pack()
    w.place(relx=0.5, rely=0.3, anchor='center')

    tkinter.Label(root, text=label).place(
        relx=0.4, rely=0.4, anchor='center')

    m = tkinter.OptionMenu(root, value, *data)
    m.pack()
    m.place(relx=0.6, rely=0.4, anchor='center')

    button = tkinter.Button(root, text='Ok', width=25, command=root.destroy).place(
        relx=0.5, rely=0.5, anchor='center')
    tkinter.mainloop()

    return value.get()
