# Excel-Python
A program to build SQL scripts from data in an Excel spreadsheet.

## Prerequisites

Openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.

Install openpyxl using pip. It is advisable to do this in a Python virtualenv without system packages:

```bash
pip install openpyxl
```
<br/>

A SQL Server can be used in this program in order to get SQL table info to build a template from which data can be entered and scripts be made. SQL Server is only necessary if you would like to build a template from an existing table. If a user desires, they may use the example.xlsx included in this repository in order to manually build an Excel spreadsheet to be used by the script building mode of this program.

If you would like to use the template building mode of this program, pyodbc software must be installed.

Pyodbc is a Python DB API 2 module for ODBC. 

Install pyodbc using pip:

```bash
pip install pyodbc
```

## Usage

To start the program open the excel.py file using python.

```bash
python excel.py
```

The program has two main run modes. In order to generate scripts from an excel file using this program, the Excel file you're reading from has to contain certain information and be formatted in a certain way. One of the functions of this program allows the user to create an Excel template in which they can deposit their data to the be used to write scripts (the other mode of the program). 

Upon starting the Excel Python program the user will be prompted to choose whether they'd like to "Build Excel template" or "Write SQL scripts". 

#### Build Excel template

The user will be prompted to enter a SQL Server name, database name, and a table name. 

If all of this info is accepted the user will be prompted to choose the types of scripts this template will be used to build (insert, update, delete, or select). 

Based on the type of script that is choosen, the user will then be prompted to choose which data will be included in the include statement and/or where statement of the SQL scripts that will be generated (Keep in mind that these can be changed in the Excel spreadsheet once the template is made). Some of the values may be greyed out and unable to be changed. This is due to the fact that these values may be necessary or forbidden to be added depending on the script and type of data. For example, you cannot chose a value to insert into SQL on an identity column and you must choose a value to insert if the column does not allow nulls.

Once all of the include/where statements have been decided on, the user will be asked to choose a file to save the template to. The template should be saved as an .xlsx file. 

#### Write SQL scripts

The user will first be shown a window explaining the proper formatting of the Excel spreadsheet they plan to use to create their scripts. They will then be asked to select the Excel spreadsheet that is formatted to be compatable with this program and is populated with data that they desire to be turned into scripts. At this point the spreadsheet needs to be validated by the program to ensure that the program will be able to write scripts from the data contained in the spreadsheet. 

There are two ways that a user can choose to validate their spreadsheet. The first is to connect to a SQL Server instance and database (preferred method) which will compare the design of the table specified in the Excel spreadsheet directly with the one in the database. The other way to validate the spreadsheet is a generic validation which will make sure that scripts can be written but does not guarantee that the design of the scripts match the design of the table they are being written for. If a spreadsheet passes validation, scripts will be generated for it. If it fails validation, windows will pop up explaining why validation failed and the program will close. The user should fix the problems in their spreadsheet then run the program again.

Once the spreadsheet is validated and scripts have been written, the user will be asked to choose a file to save the scripts to. The scripts should be saved as an .xlsx file. 

#### Validate Excel spreadsheet

The user will first be shown a window explaining the proper formatting of the Excel spreadsheet they plan to use to create their scripts. They will then be asked to select the Excel spreadsheet that is formatted to be compatable with this program and is populated with data that they desire to be turned into scripts.

There are two ways that a user can choose to validate their spreadsheet. The first is to connect to a SQL Server instance and database (preferred method) which will compare the design of the table specified in the Excel spreadsheet directly with the one in the database. The other way to validate the spreadsheet is a generic validation which will make sure that scripts can be written but does not guarantee that the design of the scripts match the design of the table they are being written for. If a spreadsheet passes validation, the user will be notified. If it fails validation, windows will pop up explaining why validation failed and the program will close. The user should fix the problems in their spreadsheet then run the program again.

If a user is choosing to validate more than once spreadsheet in a workbook and one or more sheets pass and one or more sheets fails, they will be greeted with a caution box that warns them to write scripts with care to ensure no mistakes are made.

## Authors
Matt Saffert
