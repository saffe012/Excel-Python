'''
Module of 'excel.py' that contains GUI functions.
Matt Saffert
12-15-2020
'''

import tkinter
from excel_constants import *
from tkinter import filedialog as tkFileDialog
import pandas as pd


def displayExcelFormatInstructions():
    '''Creates a tkinter pop-up box that explains the formatting of the excel
    spreadsheet that will be read into the program

    :return: NONE
    '''

    root = tkinter.Tk()
    addQuitMenuButton(root)
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


def saveToExcel(workbook):
    '''Saves the workbook to an Excel file.

    :param1 workbook: dict
    '''
    output_string = "Select/create the filename of Excel workbook you'd like to save/write to: "
    createPopUpBox(
        output_string)  # tkinter dialog box

    file = tkinter.Tk()
    # opens file explorer so user can choose file to write to
    file.filename = tkFileDialog.asksaveasfilename(
        initialdir="C:/", title="Select/create file to save/write to", defaultextension=".xlsx")
    # saves new workbook with generated scripts to a user selected file
    with pd.ExcelWriter(file.filename) as writer:
        for worksheet in workbook:
            workbook[worksheet].to_excel(
                writer, sheet_name=worksheet, header=False, index=False)
    file.destroy()

    output_string = "Scripts saved to: '" + \
        str(file.filename) + "'"
    createPopUpBox(
        output_string)  # tkinter dialog box


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

    return program_mode.get()


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
