'''
Module of 'excel.py' that handles functions related to the validation of
Excel Workbooks to be used to write SQL scripts.
Matt Saffert
1-20-2020
'''

import excel_global


def validationMode():
    '''Starts the validation mode of the application.

    :return: NONE
    '''

    excel_global.displayExcelFormatInstructions()  # tkinter dialog box

    output_string = "Choose the Excel workbook you'd like to validate."
    workbook = excel_global.openExcelFile(output_string)

    validate(workbook)


def validate(workbook):
    '''Cycles through worksheets in a workbook checking if they're valid.
    Alerts the user if any or all worksheets are invalid.

    :param1 workbook: dict
    '''

    validate_with_sql = excel_global.createYesNoBox(  # tkinter dialog box that asks user if they want to connect to a SQL database to validate spreadsheet
        'Would you like to validate Workbook with SQL table or generic validation?', 'SQL', 'Generic')

    any_changes = False # False if all spreadsheets fail validation
    all_sheets_okay = True # True if all spreadsheets pass validation

    for worksheet in workbook:
        # check if worksheet is is valid and if user wants to write scripts for them
        valid_template = excel_global.validWorksheet(
            workbook[worksheet], validate_with_sql, worksheet)
        all_sheets_okay = valid_template # True if spreadsheet passes validation
        if valid_template:  # only write to Excel if the Excel spreadsheet is a valid format
            output_string = "VALID. This worksheet will function properly with the 'Write SQL script' mode of this program."
            excel_global.createPopUpBox(
                output_string)  # tkinter dialog box
            any_changes = True  # changes were made and need to be saved

    if any_changes and not all_sheets_okay:  # some but not all spreadsheets in workbook pass validation
        output_string = "CAUTION. Care must be taken building scripts with this workbook because not all sheets are in a valid form."
        excel_global.createPopUpBox(output_string)
