import openpyxl
from tkinter import *
from tkinter.ttk import Radiobutton
from constants import *
from function import *
from openpyxl import load_workbook
from constants import time_now, vopros1


def clicked():
    if selected.get() == 1:
    	scan()

    if selected.get() == 2:
        def zapis():  #<-  Функция для записи данных
            file = 'Таблица учета Оплат.xlsx'
            wb = load_workbook(file)
            ws = wb['Лист1']
            number = scan2()
            ws['A' + str(number)] = number - 1
            ws['B' + str(number)] = txt.get()
            ws['C' + str(number)] = txt2.get()
            ws['D' + str(number)] = txt3.get()
            ws['E' + str(number)] = int(txt4.get())
            ws['F' + str(number)] = txt5.get()
            ws['G' + str(number)] = txt6.get()
            ws['H' + str(number)] = txt7.get()
            ws['I' + str(number)] = txt8.get()
            wb.save(file)
            wb.close()

        txt = Entry(window,width=20)
        btn = Button(window, text="ОК", command=zapis)
        lbl = Label(window, text='Введите дату счета')
        lbl.grid(column=0, row=1)
        txt.grid(column=1, row=1)
        btn.grid(column=2, row=5)
        txt2 = Entry(window,width=20)
        lbl2 = Label(window, text='Введите номер счета')
        txt2.grid(column=1, row=2)
        lbl2.grid(column=0, row=2)
        txt3 = Entry(window,width=20)
        lbl3 = Label(window, text='Введите имя организации')
        txt3.grid(column=1, row=3)
        lbl3.grid(column=0, row=3)
        txt4 = Entry(window,width=20)
        lbl4 = Label(window, text='Введите инн организации')
        txt4.grid(column=1, row=4)
        lbl4.grid(column=0, row=4)
        txt5 = Entry(window,width=20)
        lbl5 = Label(window, text='Введите всю сумму платежа')
        txt5.grid(column=1, row=5)
        lbl5.grid(column=0, row=5)
        txt6 = Entry(window,width=20)
        lbl6 = Label(window, text='Введите остаток сколько нужно доплатить')
        txt6.grid(column=1, row=6)
        lbl6.grid(column=0, row=6)
        txt7 = Entry(window,width=20)
        lbl7 = Label(window, text='Введите дату до которой нужно внести остаток суммы')
        txt7.grid(column=1, row=7)
        lbl7.grid(column=0, row=7)
        txt8 = Entry(window,width=20)
        lbl8 = Label(window, text='Введите статус Оплачен или Не оплачен')
        txt8.grid(column=1, row=8)
        lbl8.grid(column=0, row=8)

    if selected.get() == 3:
        def name():
            scan3(txt.get())

        txt = Entry(window,width=20)
        btn = Button(window, text="ОК", command=name)
        lbl = Label(window, text='Введите название  ОРГАНИЗАЦИИ')
        lbl.grid(column=0, row=4)
        txt.grid(column=1, row=4)
        btn.grid(column=2, row=4)

    if selected.get() == 4:
        def name():
            scan4(int(txt.get()))

        txt = Entry(window,width=20)
        btn = Button(window, text="ОК", command=name)
        lbl = Label(window, text='Введите ИНН')
        lbl.grid(column=0, row=4)
        txt.grid(column=1, row=4)
        btn.grid(column=2, row=4)

    if selected.get() == 5:
        def name():
            scan5(txt.get())

        txt = Entry(window,width=20)
        btn = Button(window, text="ОК", command=name)
        lbl = Label(window, text='Введите номер СЧЕТА')
        lbl.grid(column=0, row=4)
        txt.grid(column=1, row=4)
        btn.grid(column=2, row=4)
        
       
window = Tk()
window.title("Приложение СРЕДСТВА ЗАЩИТЫ")
window.geometry('700x400')
selected = IntVar()

rad1 = Radiobutton(window,text='ЗАДОЛЖНОСТЬ', value=1, variable=selected)
rad2 = Radiobutton(window,text='ЗАПИСАТЬ', value=2, variable=selected)
rad3 = Radiobutton(window,text='ОРГАНИЗАЦИЯ', value=3, variable=selected)
rad4 = Radiobutton(window,text='ИНН', value=4, variable=selected)
rad5 = Radiobutton(window,text='СЧЕТ', value=5, variable=selected)
btn = Button(window, text="ОК", command=clicked)
lbl = Label(window)
rad1.grid(column=0, row=0)
rad2.grid(column=1, row=0)
rad3.grid(column=2, row=0)
rad4.grid(column=3, row=0)
rad5.grid(column=4, row=0)
btn.grid(column=5, row=0)
lbl.grid(column=0, row=1)

window.mainloop()
