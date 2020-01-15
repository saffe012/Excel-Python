'''
Program to build SQL insert, update, select, and delete scripts from data in an excel spreadsheet
Matt Saffert
12-31-2019
'''

import openpyxl
import Tkinter
import tkFileDialog
import create_excel_template as template
import write_sql_scripts as write_scripts
import excel_global


def main():
    '''Main run function for the Excel python program. Called once on program initialization

    :return: NONE
    '''

    try:
        # gets the mode of the program that the user would like to use
        program_mode = excel_global.getProgramMode().get()

        if program_mode == 'scripts':
            write_scripts.displayExcelFormatInstructions()  # Tkinter dialog box

            output_string = "Choose the Excel workbook you'd like to make scripts for."
            excel_global.createPopUpBox(output_string)  # Tkinter dialog box

            file = Tkinter.Tk()
            # opens file explorer so user can choose file to read from
            file.filename = tkFileDialog.askopenfilename(
                initialdir="C:/", title="Select file to write scripts for")
            file.destroy()

            workbook = openpyxl.load_workbook(
                filename=file.filename, data_only=True)

            write_scripts.writeToExcel(workbook)

            output_string = "Select/create the filename of Excel workbook you'd like to save/write to: "
            excel_global.createPopUpBox(output_string)  # Tkinter dialog box

            file = Tkinter.Tk()
            # opens file explorer so user can choose file to write to
            file.filename = tkFileDialog.asksaveasfilename(
                initialdir="C:/", title="Select/create file to save/write to", defaultextension=".xlsx")
            # saves new workbook with generated scripts to a user selected file
            workbook.save(file.filename)
            file.destroy()

            output_string = "Scripts saved to: '" + str(file.filename) + "'"
            excel_global.createPopUpBox(output_string)  # Tkinter dialog box

        elif program_mode == 'template':
            sql_column_names, sql_column_types, column_is_nullable, column_is_identity, sql_table_name = excel_global.getTemplateInfo()  # Tkinter dialog boxes

            # creates new workbook to write template to
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.title = sql_table_name

            # allows user to select the type of script this template is for
            script_type = template.getTypeOfScriptFromUser(
                worksheet.title).get()  # Tkinter dialog box

            # asks user which elements from the imported table they'd like to includein their scripts
            sql_include_row, sql_where_row = template.populateClauses(
                sql_table_name, sql_column_names, column_is_nullable, column_is_identity, script_type)  # Tkinter dialog boxes

            # writes the generated template to the new Excel workbook
            template.WriteTemplateToSheet(worksheet, sql_table_name, script_type,
                                          sql_column_names, sql_column_types, sql_include_row, sql_where_row)

            file = Tkinter.Tk()
            # opens file explorer so user can choose file to write to
            file.filename = tkFileDialog.asksaveasfilename(
                initialdir="C:/", title="Select/create file to save/write to", defaultextension=".xlsx")
            # saves new workbook with generated template to a user selected file
            workbook.save(file.filename)
            file.destroy()

            output_string = "Excel template saved to: '" + \
                str(file.filename) + "'"
            excel_global.createPopUpBox(output_string)  # Tkinter dialog box

    # Error handling
    except IOError as e:
        file.destroy()
        print(repr(e))
        if e[0] == 13:
            excel_global.createPopUpBox(
                "Try closing the file you're writing to if it's open.")
            print("Try closing the file you're writing to if it's open.")
        elif e[0] == 22:
            excel_global.createPopUpBox("The file name you entered is invalid.")
            print("The file name you entered is invalid.")
    except Exception as e:
        excel_global.createPopUpBox(repr(e))
        print(repr(e))


main()
