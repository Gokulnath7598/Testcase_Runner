import psycopg2
import pyodbc

import xlwt

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



con = psycopg2.connect(database="postgres", user="postgres", password="1234", host="127.0.0.1", port="5432")

print("Database opened successfully")

# Reading an excel file using Python
import xlrd

# Give the location of the file
loc = ("Gokul_Testing.xlsx")

# To open Workbook
wb1 = xlrd.open_workbook(loc)
sheet = wb1.sheet_by_index(0)

# For row 0 and column 0
input =sheet.cell_value(1, 2)
print(input)
sheet1.write(1, 0, input)

try:
    cur = con.cursor()
    cur.execute(input)
    rows = cur.fetchall()
    print(type(rows))
    print(rows)
    sheet1.write(1, 2, str(rows))

except Exception as rows:
    print(rows)
    sheet1.write(1, 2, str(rows))


wb.save('xlwt example.xls')







"""import pyodbc
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=GOKUL;'
                      'Database=RCS;'
                      'Trusted_Connection=yes;')

print("Database opened successfully")



##conn = pyodbc.connect(server='localhost',port=1433,database='RCS')"""



server = 'GOKUL/gocoo'
db = 'RCS'
usern = 'Gokul'
pwd = 'Gokul'
tcon ='yes'
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+db+';UID='+usern+';PWD='+pwd+'')
print("Database opened successfully")


# Reading an excel file using Python
import xlrd

# Give the location of the file
loc = ("Gokul_Testing.xlsx")

# To open Workbook
sheet = wb1.sheet_by_index(1)

# For row 0 and column 0
input =sheet.cell_value(1, 2)
print(input)
sheet1.write(1, 1, input)
try:
    cur = cnxn.cursor()
    cur.execute(input)
    rows = cur.fetchall()
    print(type(rows1))
    print(rows1)
    sheet1.write(1, 3, str(rows))

except Exception as rows:
    print(rows1)
    sheet1.write(1, 3, str(rows))


if rows==rows1:
    print("Pass")
    sheet1.write(1, 4, 'pass')

else:
    print("Fail")
    sheet1.write(1, 4, 'Fail')









