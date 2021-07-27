
# Добавление в таблицу НМУ
cursor = cnxn.cursor()
for NUM in range(1, 10981):
    NMU_Code = sheet[f'A{NUM}'].value
    NMU_Name = sheet[f'B{NUM}'].value
    cursor.execute("insert into NMU (UNAME, UGUID, Code, NMUID) values ('" + NMU_Name + "', (select NEWID()), '" + NMU_Code + "', '" + str(NUM) + "')")
    cursor.commit()
