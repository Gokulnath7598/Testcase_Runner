
import pyodbc
import xlwt
import xlrd

from xlwt import Workbook

# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Postgres_Outputs')
sheet1.write(0, 0, 'PG_Input')
sheet1.write(0, 1, 'MSS_Input')
sheet1.write(0, 2, 'PG_Output')
sheet1.write(0, 3, 'MSS_Output')
sheet1.write(0, 4, 'Result')

# Give the location of the file
loc = ("Gokul_Testing.xlsx")
# To open Workbook
wb1 = xlrd.open_workbook(loc)
sheet = wb1.sheet_by_index(1)

i=1
for i in range(1,11) :
    cnxn = pyodbc.connect(
        'Driver={ODBC Driver 17 for SQL Server};''Server=GOKUL;''Database=RCS;''Trusted_connection=yes;')


    print("Database opened successfully")

    # Reading an excel file using Python

    # For row 0 and column 0
    input = sheet.cell_value(i, 2)
    print(input)
    sheet1.write(i, 1, input)
    try:
        cur = cnxn.cursor()
        cur.execute(input)
        rows1 = cur.fetchall()
        print(type(rows1))
        print(rows1)
        sheet1.write(i, 3, str(rows1))

    except Exception as rows1:
        print(rows1)
        sheet1.write(i, 3, str(rows1))



wb.save('xlwt example.xls')
