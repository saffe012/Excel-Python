'''
Module of 'excel.py' that contains global functions.
Matt Saffert
1-9-2020
'''

import Tkinter
import pyodbc
import constants as cons


def getExcelCellToInsertInto(column, row):
    '''Gets the column and row of the excel spreadsheet that the script should be inserted into.

    :param1 table: tuple
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


def getTemplateInfo():
    '''Creates a series of Tkinter dialogues that asks user to info about the
    template they are trying to create.

    :return: List[str], List[str], List[str], List[int], str
    '''

    description = "Please enter the name of the SQL Server where your database is located:"
    label = 'SQL Server name: '
    sql_server_name = createTextEntryBox(description, label).get()

    description = "Please enter the name of the database where the table you'd like to work with is located:"
    label = 'SQL database name: '
    sql_database_name = createTextEntryBox(description, label).get()

    description = "Please enter the name of the table you'd like to work with in the " + \
        sql_database_name + " database:"
    label = 'SQL table name: '
    sql_table_name = createTextEntryBox(description, label).get()

    # opens connection to specified SQL server and database
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=' + sql_server_name + ';'
                          'Database=' + sql_database_name + ';'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()

    # executes SQL script on database connection
    cursor.execute("SELECT info.COLUMN_NAME, info.DATA_TYPE, info.IS_NULLABLE, sy.is_identity FROM INFORMATION_SCHEMA.COLUMNS info, sys.columns sy WHERE info.TABLE_NAME = '" +
                   sql_table_name + "' AND sy.object_id = object_id('" + sql_table_name + "') AND sy.name = info.COLUMN_NAME;")

    # Lists used to hold the values retireved from SQL script in their corresponding indexes
    sql_column_names = []
    sql_column_types = []
    column_is_nullable = []
    column_is_identity = []

    # populate row lists with values from the serlect script
    for row in cursor:
        sql_column_names.append(row[0])
        sql_column_types.append(row[1])
        column_is_nullable.append(row[2])
        column_is_identity.append(row[3])

    return sql_column_names, sql_column_types, column_is_nullable, column_is_identity, sql_table_name


def getProgramMode():
    '''Creates a Tkinter dialog box that asks the user what mode they'd like
    the program to enter. (build an Excel template or write scripts)

    :return: instance
    '''

    root = Tkinter.Tk()
    root.title('Excel Python')
    root.geometry("500x500")
    program_mode = Tkinter.StringVar()
    program_mode.set("scripts")
    w = Tkinter.Label(
        root, text="Would you like to build an Excel template or write SQL scripts to an Excel file: ")
    w.pack()
    w.place(relx=0.5, rely=0.1, anchor='center')
    Tkinter.Radiobutton(root, text='Build Excel template', variable=program_mode,
                        value='template').place(relx=0.5, rely=0.4, anchor='center')
    Tkinter.Radiobutton(root, text='Write SQL scripts', variable=program_mode,
                        value='scripts').place(relx=0.5, rely=0.5, anchor='center')
    button = Tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.6, anchor='center')
    Tkinter.mainloop()

    return program_mode


def createPopUpBox(output_string):
    '''Creates a Tkinter pop-up box that displays whatever test is input with an "Ok" button
    to acknowledge info/close window

    :param1 output_string: str

    :return: NONE
    '''

    root = Tkinter.Tk()
    root.title('Excel Python')
    root.geometry("450x150")
    w = Tkinter.Label(root, text=output_string)
    w.pack()
    w.place(relx=0.5, rely=0.2, anchor='center')
    button = Tkinter.Button(root, text='Ok', width=25, command=root.destroy).place(
        relx=0.5, rely=0.5, anchor='center')
    Tkinter.mainloop()


def createTextEntryBox(description, label):
    '''Creates a Tkinter dialog box that asks the user to enter the requested
    information in a text box.

    :param1 description: str
    :param1 label: str

    :return: instance
    '''

    root = Tkinter.Tk()
    root.title('Excel Python')
    root.geometry("600x400")
    entry_value = Tkinter.StringVar()
    w = Tkinter.Label(root, text=description)
    w.pack()
    w.place(relx=0.5, rely=0.3, anchor='center')
    Tkinter.Label(root, text=label).place(
        relx=0.4, rely=0.4, anchor='center')
    e1 = Tkinter.Entry(root, textvariable=entry_value)
    e1.place(relx=0.6, rely=0.4, anchor='center')
    button = Tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.5, anchor='center')
    Tkinter.mainloop()

    return entry_value
