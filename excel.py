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
from excel_constants import *


def main():
    '''Main run function for the Excel python program. Called once on program initialization

    :return: NONE
    '''

    # try:
    # gets the mode of the program that the user would like to use
    program_mode = excel_global.getProgramMode()

    if program_mode == 'scripts':
        write_scripts.writeMode()

    elif program_mode == 'template':
        template.templateMode()

    elif program_mode == 'validate':
        validate.validationMode()


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
