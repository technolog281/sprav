import pyodbc
import openpyxl

wb = openpyxl.load_workbook('C:\\SPRAV\\nmu.xlsx')
sheet = wb['1']

server = 'localhost\SQLMIS'
database = 'research'
username = 'sa'
password = 'ApacheServer1390'
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='
                      + server + ';DATABASE=' + database + ';UID='
                      + username + ';PWD=' + password)

cursor = cnxn.cursor()
for NUM in range(1, 10981):
    NMU_Code = sheet[f'A{NUM}'].value
    NMU_Name = sheet[f'B{NUM}'].value
    cursor.execute("insert into FSIDI (FSIDIName, UGUID, CodeNMU, CodeLIS, FSIDIID) values ('"
                   + Fullname + "', (select NEWID()), '" + NMU_Code + "', '" + str(NUM) + "')")
    cursor.commit()
