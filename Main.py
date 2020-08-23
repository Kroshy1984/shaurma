import sqlite3
import tkinter
from tkinter.ttk import Treeview
import xlrd
from tkinter import filedialog

class sayrma():
    def exel_load(self,filename):
        book = xlrd.open_workbook(filename)
        List = book.sheet_by_index(0)
        for row in range(1, List.nrows):
            name2 = List.cell(row, 0)
            name3 = List.cell(row, 3)
            name2 = str(name2)
            name2 = name2[6:-1]
            name3 = str(name3)
            name3 = name3[7:]
            sql = "SELECT * FROM Pozition WHERE Name like ?"
            con = sqlite3.connect('Store.db')
            cur = con.cursor()
            for row in cur.execute("SELECT * FROM Pozition where Name like ?", (name2,)):
                print(row)

            for row in cur.execute(sql, (name2,)):

                c = row[2] - float(name3)
                sql2 = "UPDATE Pozition SET Total = ? WHERE Name like ?"
                cur.execute(sql2, (c, name2))
                con.commit()

            for row in cur.execute("SELECT * FROM Pozition where Name like ?", (name2,)):
                print(row)
                self.table.delete(*self.table.get_children())
                self.sqlite_create()

    def sqlite_create(self):
        con = sqlite3.connect('Store.db')
        cur = con.cursor()
        spt = 0
        for row in cur.execute('SELECT * FROM Pozition'):
            self.table.insert("", "end", text=str(spt), values=row)
            spt += 1
        con.close()

    def insert2(self):
        con = sqlite3.connect('Store.db')
        cur = con.cursor()
        cur.execute(''' INSERT INTO Pozition VALUES(?, ?, ?, ?, ?, ?); ''', (self.a, self.b, self.c, self.d, self.e, self.z))
        con.commit()
        con.close()

    def delete(self):
        select = self.table.focus()
        select2 = self.table.item(select, option="values")
        a = select2[0]
        con = sqlite3.connect('Store.db')
        cur = con.cursor()
        cur.execute(''' DELETE FROM Pozition WHERE Name = ?;''', (a,))
        con.commit()
        con.close()
        self.table.delete(*self.table.get_children())
        self.sqlite_create()

    def edit(self):
        self.edit = tkinter.Toplevel(self.window)
        self.edit.geometry("500x350")
        self.edit.title("Добавить")
        lable = tkinter.Label(self.edit, text="Название позиции", fg="Black")
        lable.place(x=30, y=50)
        lable1 = tkinter.Label(self.edit, text="Единица учета", fg="Black")
        lable1.place(x=30, y=70)
        lable2 = tkinter.Label(self.edit, text="Количество единиц", fg="Black")
        lable2.place(x=30, y=90)
        lable3 = tkinter.Label(self.edit, text="Цена за единицу", fg="Black")
        lable3.place(x=30, y=110)
        lable4 = tkinter.Label(self.edit, text="Цена реализации за единицу", fg="Black")
        lable4.place(x=30, y=130)
        lable5 = tkinter.Label(self.edit, text="Комментарий", fg="Black")
        lable5.place(x=30, y=150)

        self.entry = tkinter.Entry(self.edit, textvariable="")
        self.entry.place(x=230, y=50)
        self.entry1 = tkinter.Entry(self.edit, textvariable="")
        self.entry1.place(x=230, y=70)
        self.entry2 = tkinter.Entry(self.edit, textvariable="")
        self.entry2.place(x=230, y=90)
        self.entry3 = tkinter.Entry(self.edit, textvariable="")
        self.entry3.place(x=230, y=110)
        self.entry4 = tkinter.Entry(self.edit, textvariable="")
        self.entry4.place(x=230, y=130)
        self.entry5 = tkinter.Entry(self.edit, textvariable="")
        self.entry5.place(x=230, y=130)
        self.entry6 = tkinter.Entry(self.edit, textvariable="")
        self.entry6.place(x=230, y=150)

        btn = tkinter.Button(self.edit, text="Отменить", bg="lightgreen", fg="Black", command=self.edit.destroy)
        btn.place(x=30, y=190)
        btn2 = tkinter.Button(self.edit, text="Внести", bg="lightgreen", fg="Black", command=self.insert)
        btn2.place(x=110, y=190)

    def edit2(self):
        select = self.table.focus()
        select2 = self.table.item(select, option="values")
        a = select2[0]
        b = select2[1]
        c = select2[2]
        d = select2[3]
        e = select2[4]
        z = select2[5]

        self.edit = tkinter.Toplevel(self.window)
        self.edit.geometry("500x350")
        self.edit.title("Редактировать")
        lable = tkinter.Label(self.edit, text="Название позиции", fg="Black")
        lable.place(x=30, y=50)
        lable1 = tkinter.Label(self.edit, text="Единица учета", fg="Black")
        lable1.place(x=30, y=70)
        lable2 = tkinter.Label(self.edit, text="Количество единиц", fg="Black")
        lable2.place(x=30, y=90)
        lable3 = tkinter.Label(self.edit, text="Цена за единицу", fg="Black")
        lable3.place(x=30, y=110)
        lable4 = tkinter.Label(self.edit, text="Цена реализации за единицу", fg="Black")
        lable4.place(x=30, y=130)
        lable5 = tkinter.Label(self.edit, text="Комментарий", fg="Black")
        lable5.place(x=30, y=150)

        self.entry = tkinter.Entry(self.edit, textvariable="")
        self.entry.place(x=230, y=50)
        self.entry.insert(0, a)
        self.entry1 = tkinter.Entry(self.edit, textvariable="")
        self.entry1.place(x=230, y=70)
        self.entry1.insert(0, b)
        self.entry2 = tkinter.Entry(self.edit, textvariable="")
        self.entry2.place(x=230, y=90)
        self.entry2.insert(0, c)
        self.entry3 = tkinter.Entry(self.edit, textvariable="")
        self.entry3.place(x=230, y=110)
        self.entry3.insert(0, d)
        self.entry4 = tkinter.Entry(self.edit, textvariable="")
        self.entry4.place(x=230, y=130)
        self.entry4.insert(0, e)
        self.entry5 = tkinter.Entry(self.edit, textvariable="")
        self.entry5.place(x=230, y=150)
        self.entry5.insert(0, z)

        self.name = a

        btn = tkinter.Button(self.edit, text="Отменить", bg="lightgreen", fg="Black", command=self.edit.destroy)
        btn.place(x=30, y=190)
        btn2 = tkinter.Button(self.edit, text="Внести", bg="lightgreen", fg="Black", command=self.editor)
        btn2.place(x=110, y=190)

    def editor(self):
        self.a = self.entry.get()
        self.b = self.entry1.get()
        self.c = self.entry2.get()
        self.d = self.entry3.get()
        self.e = self.entry4.get()
        self.z = self.entry5.get()

        con = sqlite3.connect('Store.db')
        cur = con.cursor()
        cur.execute('''UPDATE Pozition SET Name=?, Units=?, Total=?, Cost=?, RURTorg=?, Komment=? WHERE Name=?''',
                    (self.a, self.b, self.c, self.d, self.e, self.z, self.name))
        con.commit()
        con.close()
        self.edit.destroy()
        self.table.delete(*self.table.get_children())
        self.sqlite_create()

    def insert(self):
        self.a = self.entry.get()
        self.b = self.entry1.get()
        self.c = self.entry2.get()
        self.d = self.entry3.get()
        self.e = self.entry4.get()
        self.z = self.entry5.get()
        self.insert2()
        self.table.delete(*self.table.get_children())
        self.edit.destroy()
        self.sqlite_create()

    def __init__(self):

        self.window = tkinter.Tk()
        self.window.geometry("1500x700")
        self.window.title("СКЛАД"),
        self.table = Treeview(self.window, columns=("Name", "Units", "Total", "Cost", "RURTorg", "Komment"), height=15)
        self.table.column("Name", width=350, anchor=tkinter.W)
        self.table.column("Units", width=130, anchor=tkinter.W)
        self.table.column("Total", width=130, anchor=tkinter.W)
        self.table.column("Cost", width=120, anchor=tkinter.W)
        self.table.column("RURTorg", width=120, anchor=tkinter.W)
        self.table.column("Komment", width=350, anchor=tkinter.W)
        self.table['show'] = 'headings'
        self.table.heading('Name', text='Наименование товара')
        self.table.heading('Units', text='Единица измерения')
        self.table.heading('Total', text='Количество товара')
        self.table.heading('Cost', text='Цена закупки')
        self.table.heading('RURTorg', text='Цена продажи')
        self.table.heading('Komment', text='Коментарии')
        self.sqlite_create()

        self.table.place(x=5, y=5)
        btn = tkinter.Button(self.window, text="Добавить", bg="lightgreen", fg="Black", command=self.edit)
        btn.place(x=1250, y=50)
        btn2 = tkinter.Button(self.window, text="Редактировать", bg="lightgreen", fg="Black", command=self.edit2)
        btn2.place(x=1250, y=80)
        btn3 = tkinter.Button(self.window, text="Удалить", bg="red", fg="Black", command=self.delete)
        btn3.place(x=1250, y=110)

        btn4 = tkinter.Button(self.window, text="Загрузить файл с данными о продажах", bg = "Lightblue", fg= "Black", command=self.OpenFile)
        btn4.place(x=250, y=550)

        self.window.mainloop()

    def OpenFile(self):
        while True:
            try:
                filename = tkinter.filedialog.askopenfilename()
                self.exel_load(filename)
                break
            except xlrd.biffh.XLRDError:
                self.errormessage(filename)


    def errormessage(self,filename):
        self.errorwindow = tkinter.Toplevel(self.window)
        self.errorwindow.geometry("700x150+500+300")
        self.errorwindow.title("Ошибка!!!!")
        self.errorwindow.lift()
        self.errorwindow.focus()
        self.errorwindow.attributes('-topmost', True)
        self.errorwindow.update()
        lable = tkinter.Label(self.errorwindow, text=f"Ошибка открытия файла {filename}!", fg="Black")
        lable.place(x=30, y=10)
        lable1 = tkinter.Label(self.errorwindow, text="Файл должен быть Excel!!", fg="Red")
        lable1.place(x=30, y=40)
        btn3 = tkinter.Button(self.errorwindow, text="Удалить", bg="red", fg="Black")
        btn3.place(x=45, y=110)

window = sayrma()

# нужно создать поиск по базе по ключевым словам
# В основном окне базы надо что позиуии были не в центре а с левого края (выравнивание) 


# Нужно еще создать одну базу по технологической карте
# И третья база по реализациям - третья база работает на основе первых двух (берет технологические расчеты и редактирует
# первую базу - уменьшая остатки