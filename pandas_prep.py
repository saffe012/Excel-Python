import pandas as pd
workbook = pd.read_excel(
    r'/home/matt/Projects/Excel-Python/data/example.xlsx', header=None, sheet_name=None)

for worksheet in workbook:
    # print(worksheet)
    workbook[worksheet] = workbook[worksheet].rename(index={0: "info"})
    workbook[worksheet] = workbook[worksheet].rename(index={1: "names"})
    workbook[worksheet] = workbook[worksheet].rename(index={2: "types"})
    workbook[worksheet] = workbook[worksheet].rename(index={3: "include"})
    workbook[worksheet] = workbook[worksheet].rename(index={4: "where"})

    for i in range(5, len(workbook[worksheet])):
        workbook[worksheet] = workbook[worksheet].rename(index={i: (i - 5)})

    print(type(workbook[worksheet]))

    print(type(workbook[worksheet].iloc[0][0]))
    # print(workbook[worksheet].loc['names'])
    script_type = workbook[worksheet].loc['info'][0]
    # print(script_type)
