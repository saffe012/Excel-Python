'''
Module of 'excel.py' that contains GUI functions.
Matt Saffert
12-15-2020
'''

import tkinter
from excel_constants import *
from tkinter import filedialog as tkFileDialog
import pandas as pd


def saveToExcel(workbook):
    '''Saves the workbook to an Excel file.

    :param1 workbook: dict

    :return: NONE
    '''

    output_string = "Select/create the filename of Excel workbook you'd like to save/write to: "
    createPopUpBox(output_string)  # tkinter dialog box

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

    output_string = "Scripts saved to: '" + str(file.filename) + "'"

    createPopUpBox(output_string)  # tkinter dialog box


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

    workbook = reformatExcelInput(workbook)

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
    description = "Are you sure you want to close the program?"
    root = generateWindow("450x150", description, relx=0.5, rely=0.2)

    str_len = len(description)
    text_height = (str_len // 35) + 1

    tkinter.Button(root, text='Yes', width=15, command=lambda: quit()).place(
        relx=0.35, rely=0.8, anchor='center')
    tkinter.Button(root, text='No', width=15, command=root.destroy).place(
        relx=0.65, rely=0.8, anchor='center')
    tkinter.mainloop()


def createTwoChoiceBox(description, label1, label2, dimensions="500x500", additional_box=(False,'')):
    '''Creates a tkinter pop-up box that gives the user a choice between 2 options

    :param1 description: str
    :param2 label1: str
    :param3 label2: str
    :param4 dimensions: str
    :param5 additional_box: tuple (bool, str)

    :return: str, tkinter.IntVar
    '''

    root = generateWindow(dimensions, description, relx=0.5, rely=0.1)

    program_mode = tkinter.StringVar()
    program_mode.set(label1)

    tkinter.Radiobutton(root, text=label1, variable=program_mode,
                        value=label1).place(relx=0.5, rely=0.4, anchor='center')
    tkinter.Radiobutton(root, text=label2, variable=program_mode,
                        value=label2).place(relx=0.5, rely=0.5, anchor='center')
    additional_box_val = tkinter.IntVar()
    if additional_box[0]:
        createCheckBox(root, additional_box[1], additional_box_val, 0.5, 0.6, select=False)
    tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.7, anchor='center')
    tkinter.mainloop()

    return program_mode.get(), additional_box_val


def getProgramMode():
    '''Creates a tkinter dialog box that asks the user what mode they'd like
    the program to enter. (build an Excel template or write scripts)

    :return: instance
    '''

    description = "Would you like to build an Excel template or write SQL scripts to an Excel file: "
    root = generateWindow("500x500", description, relx=0.5, rely=0.1)

    program_mode = tkinter.StringVar()
    program_mode.set("scripts")

    tkinter.Radiobutton(root, text='Check if workbook is valid for writing scripts', variable=program_mode,
                        value='validate').place(relx=0.5, rely=0.3, anchor='center')
    tkinter.Radiobutton(root, text='Build Excel template', variable=program_mode,
                        value='template').place(relx=0.5, rely=0.4, anchor='center')
    tkinter.Radiobutton(root, text='Write SQL scripts', variable=program_mode,
                        value='scripts').place(relx=0.5, rely=0.5, anchor='center')
    tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
        relx=0.5, rely=0.6, anchor='center')
    tkinter.mainloop()

    return program_mode.get()


def createPopUpBox(description, dimensions="450x150"):
    '''Creates a tkinter pop-up box that displays whatever text is input with an "Ok" button
    to acknowledge info/close window

    :param1 description: str
    :param2 dimensions: str
    '''

    root = generateWindow(dimensions, description, relx=0.5, rely=0.2)

    tkinter.Button(root, text='Ok', width=25, command=root.destroy).place(
        relx=0.5, rely=0.8, anchor='center')
    tkinter.mainloop()


def createTextEntryBox(description, label):
    '''Creates a tkinter dialog box that asks the user to enter the requested
    information in a text box.

    :param1 description: str
    :param2 label: str

    :return: instance
    '''

    root = generateWindow("600x400", description, relx=0.5, rely=0.3)

    entry_value = tkinter.StringVar()

    tkinter.Label(root, text=label).place(
        relx=0.4, rely=0.4, anchor='center')
    e1 = tkinter.Entry(root, textvariable=entry_value)
    e1.place(relx=0.6, rely=0.4, anchor='center')
    tkinter.Button(root, text='Next', width=25, command=root.destroy).place(
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

    root = generateWindow("500x500", description, relx=0.5, rely=0.3)

    value = tkinter.StringVar(root)
    value.set(data[0])  # default value

    tkinter.Label(root, text=label).place(
        relx=0.4, rely=0.4, anchor='center')

    m = tkinter.OptionMenu(root, value, *data)
    m.pack()
    m.place(relx=0.6, rely=0.4, anchor='center')

    tkinter.Button(root, text='Ok', width=25, command=root.destroy).place(
        relx=0.5, rely=0.5, anchor='center')
    tkinter.mainloop()

    return value.get()


def generateBox(dimensions):
    '''Creates a base tkinter dialog box to build on.

    :param1 dimensions: str

    :return: ?
    '''

    root = tkinter.Tk()
    addQuitMenuButton(root)
    root.title('SQL Generator')
    root.geometry(dimensions)

    return root


def addLabelToBox(root, description, relx=0.5, rely=0.2, anchor='center'):
    '''Inserts text into a tkinter dialog box.

    :param1 root: ?
    :param2 description: str
    :param3 relx: float
    :param4 rely: float
    :param5 anchor: str

    :return: NONE
    '''

    w = tkinter.Label(root, text=description)
    w.pack()
    w.place(relx=relx, rely=rely, anchor=anchor)


def reformatExcelInput(workbook):
    '''Reformats the inputed Excel data to work with program.

    :param1 workbook: dict

    :return: dict
    '''

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


def generateWindow(dimensions, description, relx=0.5, rely=0.2, anchor='center'):
    '''Inserts text into a tkinter dialog box.

    :param1 dimensions: str
    :param2 description: str
    :param3 relx: float
    :param4 rely: float
    :param5 anchor: str

    :return: ?
    '''

    root = generateBox(dimensions)
    addLabelToBox(root, description, relx, rely, anchor)

    return root


def createCheckBox(root, column_name, include_value, x_spacing, y_spacing, select=False, disable='normal'):
    '''Inserts text into a tkinter dialog box.

    :param1 root: ?
    :param2 column_name: str
    :param3 include_value: str
    :param4 select: bool
    :param5 disable: bool
    :param6 x_spacing: float
    :param7 y_soacing: float

    :return: NONE
    '''

    b = tkinter.Checkbutton(
        root, text=column_name, variable=include_value, state=disable)
    if select:
        b.select()
    else:
        b.deselect()
    b.place(relx=x_spacing, rely=y_spacing, anchor='center')
