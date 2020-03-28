'''
Program to build SQL insert, update, select, and delete scripts from data in an excel spreadsheet
Matt Saffert
12-31-2019
'''

import create_excel_template as template
import write_sql_scripts as write_scripts
import validate_workbook as validate
import excel_global
import sys
import os
import pandas as pd


def main():
    '''Main run function for the Excel python program. Called once on program initialization

    :return: NONE
    '''

    # try:
    # gets the mode of the program that the user would like to use
    program_mode = excel_global.getProgramMode().get()

    if program_mode == 'scripts':
        write_scripts.displayExcelFormatInstructions()  # tkinter dialog box

        output_string = "Choose the Excel workbook you'd like to make scripts for."
        workbook = excel_global.openExcelFile(output_string)

        validate_with_sql = excel_global.createYesNoBox(
            'Would you like to validate Workbook with SQL table or generic validation?', 'SQL', 'Generic')

        write_to_sql = 'SQL'
        write_to_excel = 'Excel'
        description = 'Would you like to write the sql scripts to a ".sql" file or to an Excel spreadsheet?'
        write_to = excel_global.createYesNoBox(  # write scripts to new SQL or Excel file
            description, write_to_sql, write_to_excel)

        if write_to == 'SQL':
            save_file = write_scripts.writeToSQL(workbook, validate_with_sql)
        elif write_to == 'Excel':
            save_file = write_scripts.writeToExcel(workbook, validate_with_sql)

        if save_file == '':  # no scripts were written because there were no valid worksheets
            output_string = "No files were changed. Closing program."
            excel_global.createPopUpBox(
                output_string)  # tkinter dialog box

    elif program_mode == 'template':
        template_type = template.getTemplateType()
        #TODO: Test making template from sql
        if template_type == 'from_table':  # generates an Excel template from a SQL database
            sql_column_names, sql_column_types, column_is_nullable, column_is_identity, sql_table_name = template.getTemplateInfo()  # tkinter dialog boxes
            workbook = {sql_table_name: pd.DataFrame()}
            worksheet = workbook[sql_table_name]

            # allows user to select the type of script this template is for
            script_type = template.getTypeOfScriptFromUser(
                sql_table_name).get()  # tkinter dialog box

            # asks user which elements from the imported table they'd like to include in their scripts
            sql_include_row, sql_where_row, disable_include_change = template.populateClauses(
                sql_table_name, sql_column_names, column_is_nullable, column_is_identity, script_type)  # tkinter dialog boxes

            # writes the generated template to the new Excel workbook
            template.WriteTemplateToSheet(worksheet, sql_column_names, sql_column_types, sql_include_row, sql_where_row, disable_include_change)
        elif template_type == 'generic':  # generates a generic template with default table data
            generic_data = cons.GENERIC_TEMPLATE # dictionary filled with generic data to build template
            worksheet = pd.DataFrame(data=generic_data)
            workbook = {'IOChannels': worksheet}
        else:
            excel_global.closeProgram()

        write_scripts.saveToExcel(workbook)

    elif program_mode == 'validate':
        write_scripts.displayExcelFormatInstructions()  # tkinter dialog box

        output_string = "Choose the Excel workbook you'd like to validate."
        workbook = excel_global.openExcelFile(output_string)

        validate.validate(workbook)

'''
    except Exception as e:
        excel_global.createErrorBox(repr(e))
        print(repr(e))
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        sys.exit()
'''

main()
