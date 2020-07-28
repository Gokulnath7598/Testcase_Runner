import psycopg2
import pyodbc
import xlwt
import xlrd

from xlwt import Workbook

#Output sheet
wb = Workbook()

sheet1 = wb.add_sheet('Postgres&MSSQL_Outputs')
sheet1.write(0, 0, 'PG_Input')
sheet1.write(0, 1, 'MSS_Input')
sheet1.write(0, 2, 'PG_Output')
sheet1.write(0, 3, 'MSS_Output')
sheet1.write(0, 4, 'Result')


#Input Sheet

loc = ("Gokul_Testing.xlsx")


inputXSL = xlrd.open_workbook(loc)
inputsheet1 = inputXSL.sheet_by_index(0)
inputsheet2 = inputXSL.sheet_by_index(1)
j=inputsheet1.nrows

i=1
for i in range(1,j) :
    pop=''
    mop=''
    print(pop)
    print(mop)

    con = psycopg2.connect(database="postgres", user="postgres", password="1234", host="127.0.0.1", port="5432")
    input = inputsheet1.cell_value(i, 2)
    ##print(input)
    sheet1.write(i, 0, input)

    try:
        cur = con.cursor()
        cur.execute(input)
        rows = cur.fetchall()
        ##print(rows)
        sheet1.write(i, 2, str(rows))
        pop=rows


    except Exception as rows:
        ##print(rows)
        sheet1.write(i, 2, str(rows))
        pop = rows

    cnxn = pyodbc.connect(
        'Driver={ODBC Driver 17 for SQL Server};''Server=GOKUL;''Database=RCS;''Trusted_connection=yes;')





    input = inputsheet2.cell_value(i, 2)
    ##print(input)
    sheet1.write(i, 1, input)
    try:
        cur1 = cnxn.cursor()
        cur1.execute(input)
        rows1 = cur1.fetchall()
        ##print(rows1)
        sheet1.write(i, 3, str(rows1))
        mop=rows1

    except Exception as rows1:
        ##print(rows1)
        sheet1.write(i, 3, str(rows1))
        mop = rows1

    print(pop)
    print(mop)

    if pop == mop:
        print("Pass")
        sheet1.write(i, 4, 'pass')

    else:
        print("Fail")
        sheet1.write(i, 4, 'Fail')




wb.save('xlwt example.xls')





