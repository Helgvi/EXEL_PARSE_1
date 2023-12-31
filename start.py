from tkinter import Tk, Button, Label, messagebox

from tkinter.filedialog import askopenfile

import sys
import parse_xls


INFO = (
    'Программа поиска и сопоставления кодов 1С '
    'с исходными данными накладной поставщика ТД Восход.')
INFO_1 = '\n 1. Выберите файл накладной ТД Восход для начала работы программы'
INFO_ENDING = (
    '2. После обработки будет создан Exel файл накладной с кодами 1С. \n'
    'По умолчанию, файл будет сохранен в папке c:\Intake на вашем компьютере.')

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
Myl1.pack()
text_unit_2 = Label(
    myroot,
    text=INFO_ENDING,
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


def mydisplay(massage):
    """Окно сообщений"""
    messagebox.showinfo("Сообщение", massage)
    sys.exit()


text_unit_2.pack()
myroot.geometry("800x400")

mybtn1 = Button(
    myroot,
    text="Выбрать файл накладной .xls",
    command=myopen_file
    )
mybtn1.pack(pady=10)
myroot.mainloop()
