import sqlite3

import xlrd

import xlwt

ANSWER = 'Кода нет!'
TITLE_LIST = [
    'Идентификатор',
    'Код в 1C',
    'Артикул произв',
    'Наименование',
    'Количество',
    'Цена',
    'Сумма'
    ]

PATH_EXTAKE = 'c:/Intake/{}.xls'
INFO = 'Программа поиска и сопоставления кодов 1С с исходными данными накладной поставщика ТД Восход'


style_string = 'font: bold on; align: wrap 1; borders: top 2, right 2, bottom 2, left 2'
style_string1 = 'font: bold off; align: wrap 1; border: top 0x1, right 1, bottom 1, left 1'
style_string2 = 'font: bold on; align: wrap 1; border: top 0x1, right 1, bottom 1, left 1'
style = xlwt.easyxf(style_string)
style1 = xlwt.easyxf(style_string1, num_format_str='0')
style2 = xlwt.easyxf(style_string1)
style3 = xlwt.easyxf(style_string2)


Vcode = list()
otvet = list()
code_list = list()
art_list = list()
name_list = list()
quent_list = list()
price_list = list()
amount_list = list()
after_action_list = list()
titles_list = list()


def bild_list():
    "Список всех кодов из базы данных"
    con = sqlite3.connect('db.sqlite')
    cur = con.cursor()
    sql = 'SELECT Vcode FROM goods;'
    cur.execute(sql)
    for result in cur:
        Vcode.append(result[0])
    con.commit()
    con.close()
    return Vcode


def return_1C_code(code_list):
    "Проверка наличия кода в списке из базы данных"
    con = sqlite3.connect('db.sqlite')
    cur = con.cursor()
    for code in code_list:
        if Vcode.count(code) != 0:
            sql = 'SELECT Ccode FROM goods WHERE Vcode = ?;'
            for result in cur.execute(sql, [code]):
                otvet = result[0]
        else:
            otvet = ANSWER
        after_action_list.append(otvet)
    con.commit()
    con.close()
    return after_action_list


def answer_from_exel_file(path_intake):
    "Обработка входящт данных исходной таблицы"
    rb = xlrd.open_workbook(path_intake)
    print("Листов книги Exel - {0}".format(rb.nsheets))
    print("Листы файла: {0}".format(rb.sheet_names()))
    sheet = rb.sheet_by_index(0)
    num = sheet.nrows
    title = sheet.cell(0, 0).value
    totals = sheet.cell(num-1, 9).value
    for rx in range(2, num-1):
        code = sheet.row(rx)[2].value
        code_list.append(code)
        code = sheet.row(rx)[3].value
        art_list.append(code)
        code = sheet.row(rx)[5].value
        name_list.append(code)
        code = sheet.row(rx)[7].value
        quent_list.append(code)
        code = sheet.row(rx)[8].value
        price_list.append(code)
        code = sheet.row(rx)[9].value
        amount_list.append(code)
    titles_list.append(title)
    titles_list.append(totals)


def write_new_data():
    "Таблица данных постобработки"
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Накладная")
    sheet1.write(0, 0, titles_list[0])
    for num in range(7):
        sheet1.write(1, num, TITLE_LIST[num], style=style3)
    for num in range(2, len(code_list)):
        sheet1.write(num, 0, code_list[num], style=style2)
        sheet1.write(num, 1, parse_code(after_action_list[num]), style=style3)
        sheet1.write(num, 2, art_list[num], style=style1)
        sheet1.write(num, 3, name_list[num], style=style2)
        sheet1.write(num, 4, quent_list[num], style=style2)
        sheet1.write(num, 5, price_list[num], style=style2)
        sheet1.write(num, 6, amount_list[num], style=style2)
    sheet1.write(len(code_list)+1, 6, titles_list[1])
    sheet1.col(3).width = 18000
    sheet1.col(2).width = 4000
    sheet1.col(1).width = 2800
    sheet1.col(0).width = 3000
    book.save(PATH_EXTAKE.format(titles_list[0]))


def parse_code(code):
    code = str(code)
    if len(code) == 3:
        code = f'00{code}'
    elif len(code) == 4:
        code = f'0{code}'
    elif len(code) == 2:
        code = f'000{code}'
    return code


def main(path_intake):
    answer_from_exel_file(path_intake)
    bild_list()
    return_1C_code(code_list)
    print("Число записей - {}".format(len(code_list)))
    print("Число записей - {}".format(len(after_action_list)))
    write_new_data()
    print("Обработка завершена!")


if __name__ == '__main__':
    main()
