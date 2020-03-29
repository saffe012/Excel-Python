'''
Project constants for excel.py
Matt Saffert
12-31-2019
'''

import numpy as np

'''
SQL Server String data types
'''
SQL_STRING_TYPE = [
    'char',
    'varchar',
    'varchar',
    'text',
    'nchar',
    'nvarchar',
    'nvarchar',
    'ntext',
    'binary',
    'varbinary',
    'varbinary',
    'image'
]

GENERIC_TEMPLATE = {'col1': ['IOChannels', 'Id', 'int', 'include', 'where'],
                    'col2': ['insert', 'IOServersId', 'int', np.NaN, 'where'],
                    'col3': [np.NaN, 'Name', 'varchar(50)', 'include', np.NaN]}

'''
SQL Server Numeric data types
'''
SQL_NUMERIC_TYPE = [
    'bit',
    'tinyint',
    'smallint',
    'int',
    'bigint',
    'decimal',
    'numeric',
    'smallmoney',
    'money',
    'float',
    'real'
]

'''
SQL Server Date/Time data types
'''
SQL_DATETIME_TYPE = [
    'datetime',
    'datetime2',
    'smalldatetime',
    'date',
    'time',
    'datetimeoffset',
    'timestamp'
]

'''
SQL Server Date/Time data types
'''
SQL_OTHER_TYPE = [
    'sql_variant',
    'uniqueidentifier',
    'xml',
    'cursor',
    'table'
]

'''
Excel row indexes
'''
INFO_ROW = 0
COLUMN_NAMES_ROW_INDEX = 1
COLUMN_DATA_TYPE_ROW_INDEX = 2
INCLUDE_ROW_INDEX = 3
WHERE_ROW_INDEX = 4
START_OF_DATA_ROWS_INDEX = 5

'''
Info row indexes
'''
TABLE_NAME = 0
SCRIPT_TYPE = 1

'''
Types of scripts generatable by this program
'''
TYPE_OF_SCRIPTS_AVAILABLE = ['insert', 'delete', 'select', 'update']

'''
Python column index to excel letter column index
'''
LETTER_INDEX_DICT = {
    'A': 0,
    'B': 1,
    'C': 2,
    'D': 3,
    'E': 4,
    'F': 5,
    'G': 6,
    'H': 7,
    'I': 8,
    'J': 9,
    'K': 10,
    'L': 11,
    'M': 12,
    'N': 13,
    'O': 14,
    'P': 15,
    'Q': 16,
    'R': 17,
    'S': 18,
    'T': 19,
    'U': 20,
    'V': 21,
    'W': 22,
    'X': 23,
    'Y': 24,
    'Z': 25,
    'AA': 26,
    'AB': 27,
    'AC': 28,
    'AD': 29,
    'AE': 30,
    'AF': 31,
    'AG': 32,
    'AH': 33,
    'AI': 34,
    'AJ': 35,
    'AK': 36,
    'AL': 37,
    'AM': 38,
    'AN': 39,
    'AO': 40,
    'AP': 41,
    'AQ': 42,
    'AR': 43,
    'AS': 44,
    'AT': 45,
    'AU': 46,
    'AV': 47,
    'AW': 48
}
