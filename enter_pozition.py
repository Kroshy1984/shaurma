import sqlite3
import xlrd

book = xlrd.open_workbook('File.xls')
List = book.sheet_by_index(0)
for row in range(1, List.nrows):

    name2 = List.cell(row, 0)
    name3 = List.cell(row, 3)
    name2 = str(name2)
    name2 = name2[6:-1]
    name3=str(name3)
    name3=name3[7:]
    #print(name3)
    #if name2=="Лаваш":print(name2)
    sql = "SELECT * FROM Pozition WHERE Name like ?"
    con = sqlite3.connect('Store.db')
    cur = con.cursor()
    for row in cur.execute ("SELECT * FROM Pozition where Name like ?",(name2,)):
        print(row)

    for row in cur.execute(sql, (name2,)):
        #print(row)
        c = row[2] - float(name3)
        sql2 = "UPDATE Pozition SET Total = ? WHERE Name like ?"
        cur.execute(sql2, (c, name2))
        con.commit()

    for row in cur.execute ("SELECT * FROM Pozition where Name like ?",(name2,)):
        print(row)




"""a = input("введите название товара ")
sql = "SELECT * FROM Pozition WHERE Name = ?"
b = input("Введите количество продаж")
con = sqlite3.connect('Store.db')
cur = con.cursor()
for row in cur.execute(sql, (a,)):
    print(row)
    c = row[2] - int(b)
    sql2 = "UPDATE Pozition SET Total = ? WHERE Name = ?"
    cur.execute(sql2, (c, a))
for row in cur.execute(sql, (a,)):
    print(row)"""
