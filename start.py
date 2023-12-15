from tkinter import Tk, Button, Label, messagebox

from tkinter.filedialog import askopenfile

import sys
import parse_xls
import orders_gen


INFO = (
    'Алгоритм 1: предназначен для поиска и сопоставления кодов 1С '
    'с исходными данными накладной поставщика ТД Восход. в формате .xls\n'
    )
INFO_1 = (
    '1. Выберите файл накладной ТД Восход для начала работы программы \n'
    '2. После обработки будет создан Exel файл накладной с кодами 1С. \n'
    '3. По умолчанию, файл будет сохранен в папке c:\Intake на вашем  \n')

INFO_ORDERS = (
    'Алгоритм 2: анализирует складские остатков и \n'
    'формирования Exel отчет доступных для заказа позиций в формате .xls')

INFO_ORDERS_1 = (
    '\n 1. Сформируйте отчет складских остатков в 1C\n'
    '2. сохраните его в формате .xlsx \n'
    'Укажите путь к файлу отчета "кликнув" на кнопку - "Выбрать отчет"\n'
    '3. После обработки будет создан Exel файл отчета. \n'
    '4.По умолчанию, файл будет сохранен в папке c:\Intake \n'
    'на вашем компьютере. Имя файла order.xls\n'
    'Созданный ранее отчет будет перезаписан')


myroot = Tk()
myroot.title("Поисковик кода 1С")
Myl1 = Label(
    myroot,
    text=INFO,
    font=("Arial Bold", 12),
    bg="#5c3825",
    justify='right',
    fg='White',
    width=100,
)
Myl1.pack()
text_unit_1 = Label(
    myroot,
    text=INFO_1,
    font=("Arial Bold", 12),
    width=100,
)

text_unit_1.pack()


def myopen_file():
    """Выбор файла"""
    myfile = askopenfile(mode='r', filetypes=[('All Python Files', '*.xls')])
    if myfile is not None:
        msg = parse_xls.main(myfile.name)
        mydisplay(msg)
    else:
        mydisplay('Файл не выбран!')


def open_file():
    """Выбор файла"""
    myfile = askopenfile(mode='r', filetypes=[('All Python Files', '*.xlsx')])
    if myfile is not None:
        msg = orders_gen.main(myfile.name)
        mydisplay(msg)
    else:
        mydisplay('Файл не выбран!')


def mydisplay(massage):
    """Окно сообщений"""
    messagebox.showinfo("Сообщение", massage)
    sys.exit()


mybtn1 = Button(
    myroot,
    text="Выбрать файл накладной .xls",
    command=myopen_file
    )
mybtn1.pack(pady=10)


Myl1.pack()
text_unit_2 = Label(
    myroot,
    text=INFO_ORDERS,
    font=("Arial Bold", 12),
    width=100,
)

Myl1.pack()
text_unit_2 = Label(
    myroot,
    text=INFO_ORDERS_1,
    font=("Arial Bold", 12),
    width=100,
)


mybtn2 = Button(
    myroot,
    text="Выбрать отчет .xlsx",
    command=open_file
    )

text_unit_2.pack()

myroot.geometry("800x400")


mybtn2.pack(pady=10)
myroot.mainloop()
