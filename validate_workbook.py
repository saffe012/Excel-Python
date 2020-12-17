'''
Module of 'excel.py' that handles functions related to the validation of
Excel Workbooks to be used to write SQL scripts.
Matt Saffert
1-20-2020
'''

import excel_global
import global_gui as gui


def validationMode():
    '''Runs through the validation mode of the application.

    :return: NONE
    '''

    gui.displayExcelFormatInstructions()  # tkinter dialog box

    workbook = gui.openExcelFile("Choose the Excel workbook you'd like to validate.")

    validate_with_sql = gui.createYesNoBox(  # tkinter dialog box that asks user if they want to connect to a SQL database to validate spreadsheet
        'Would you like to validate Workbook with SQL table or generic validation?', 'SQL', 'Generic')

    any_valid_sheets, all_valid_sheets = validWorkbook(
        workbook, validate_with_sql)

    displayWorkbookValidationResult(any_valid_sheets, all_valid_sheets)


def validWorkbook(workbook, validate_with_sql):
    '''Cycles through worksheets in a workbook checking if they're valid.

    :param1 workbook: dict
    :param2 validate_with_sql: str
    '''

    any_valid_sheets = False  # False if all spreadsheets fail validation
    all_valid_sheets = True  # True if all spreadsheets pass validation

    for worksheet in workbook:
        # check if worksheet is is valid and if user wants to write scripts for them
        valid_worksheet = excel_global.validWorksheet(
            workbook[worksheet], validate_with_sql, worksheet)
        # True if spreadsheet passes validation
        all_valid_sheets = valid_worksheet and all_valid_sheets
        if valid_worksheet:  # only write to Excel if the Excel spreadsheet is a valid format
            output_string = "VALID. This worksheet will function properly with the 'Write SQL script' mode of this program."
            gui.createPopUpBox(
                output_string)  # tkinter dialog box
            any_valid_sheets = True  # changes were made and need to be saved

    return any_valid_sheets, all_valid_sheets


def displayWorkbookValidationResult(any_valid_sheets, all_valid_sheets):
    '''Generates a window displaying the status of the completed validation.

    :param1 any_valid_sheets: bool
    :param2 all_valid_sheets: bool
    '''

    if all_valid_sheets:
        output_string = "SUCCESS. All sheets hae been successfully validated."
    elif not any_valid_sheets:
        output_string = "FAILURE. No sheets could be successfully validated. Please review rules."
    else:  # some but not all spreadsheets in workbook pass validation
        output_string = "CAUTION. Care must be taken building scripts with this workbook because not all sheets are in a valid form."

    gui.createPopUpBox(output_string)
