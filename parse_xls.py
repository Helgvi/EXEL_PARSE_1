import sqlite3

import xlrd
import xlwt


ANSWER = 'Кода нет!'
TITLE_LIST = [
    '№ п.п',
    'Код Восход',
    'Код в 1C',
    'Арт.',
    'Наименование',
    'Кол-во',
    'Единица',
    'Цена',
    'Сумма'
    ]

PATH_EXTAKE = 'c:/Intake/{}.xls'
INFO = (
    'Программа поиска и сопоставления кодов 1С'
    'с исходными данными накладной поставщика ТД Восход')

"Table style setings"
title_doc_string = 'font: bold on, height 280;'
table_title_string = (
    'font: bold on; align: wrap 1;'
    'borders: top 2,'
    'right 2, bottom 2, left 2')
base_style_string = (
    'font: bold off; align: wrap 1;'
    'border: top 0x1, right 0x1, bottom 0x1, left 0x1')
code_staly_string = (
    'font: bold on;'
    'border: top 0x1, right 0x1, bottom 0x1, left 0x1')

title_style = xlwt.easyxf(title_doc_string)
table_title_style = xlwt.easyxf(table_title_string)
base_style = xlwt.easyxf(base_style_string, num_format_str='0')
code_style = xlwt.easyxf(code_staly_string)
quent_style = xlwt.easyxf(base_style_string, num_format_str='0')
price_style = xlwt.easyxf(base_style_string, num_format_str='#,##0.00')

Vcode = list()
otvet = list()
number = list()
code_list = list()
art_list = list()
name_list = list()
quent_list = list()
price_list = list()
size = list()
amount_list = list()
intake_list = list()
titles_list = list()
agent_list = list()


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
            sql = 'SELECT Vcode FROM goods WHERE Ccode = ?;'
            for result in cur.execute(sql, [code]):
                otvet = result
        else:
            otvet = ANSWER
        intake_list.append(otvet)
    con.commit()
    con.close()
    return intake_list


def answer_from_exel_file(path_intake):
    "Обработка входящих данных исходной таблицы"
    rb = xlrd.open_workbook(path_intake)
    print("Листов книги Exel - {0}".format(rb.nsheets))
    print("Листы файла: {0}".format(rb.sheet_names()))
    sheet = rb.sheet_by_index(0)
    num = sheet.nrows
    title = sheet.cell(0, 0).value
    agent = sheet.cell(1, 2).value
    totals = sheet.cell(num-1, 9).value
    for rx in range(4, num-1):
        code = sheet.row(rx)[0].value
        number.append(code)
        code = sheet.row(rx)[2].value
        code_list.append(code)
        code = sheet.row(rx)[3].value
        art_list.append(code)
        code = sheet.row(rx)[5].value
        name_list.append(code)
        code = sheet.row(rx)[6].value
        size.append(code)
        code = sheet.row(rx)[8].value
        price_list.append(code)
        code = sheet.row(rx)[7].value
        quent_list.append(code)
        code = sheet.row(rx)[9].value
        amount_list.append(code)
    titles_list.append(title)
    agent_list.append(agent)
    titles_list.append(totals)


def write_new_data():
    "Таблица данных постобработки"
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Накладная")
    sheet1.write(0, 0, titles_list[0], style=title_style)
    sheet1.write(1, 0, agent_list[0], style=title_style)
    for num in range(9):
        sheet1.write(4, num, TITLE_LIST[num], style=table_title_style)
    for num in range(0, len(code_list)):
        col = num+5
        sheet1.write(col, 0, number[num], style=base_style)
        sheet1.write(col, 1, code_list[num], style=base_style)
        sheet1.write(col, 2, parse_code(intake_list[num]), style=code_style)
        sheet1.write(col, 3, art_list[num], style=base_style)
        sheet1.write(col, 4, name_list[num], style=base_style)
        sheet1.write(col, 5, quent_list[num], style=quent_style)
        sheet1.write(col, 6, size[num], style=base_style)
        sheet1.write(col, 7, price_list[num], style=price_style)
        sheet1.write(col, 8, amount_list[num], style=price_style)
    sheet1.write(len(code_list)+5, 7, "Итого:", style=price_style)
    sheet1.write(len(code_list)+5, 8, titles_list[1], style=price_style)
    sheet1.col(0).width = 1500
    sheet1.col(1).width = 3000
    sheet1.col(2).width = 2300
    sheet1.col(3).width = 4000
    sheet1.col(4).width = 18000
    book.save(PATH_EXTAKE.format(titles_list[0]))


def parse_code(code):
    "Формат "
    code = str(code)
    if len(code) == 3:
        code = f'00{code}'
    elif len(code) == 4:
        code = f'0{code}'
    elif len(code) == 2:
        code = f'000{code}'
    return code


def main(path_intake):
    "Вход в программу"
    answer_from_exel_file(path_intake)
    bild_list()
    return_1C_code(code_list)
    print("Число записей - {}".format(len(code_list)))
    print("Число записей - {}".format(len(intake_list)))
    write_new_data()
    return 'Выполнено!'


if __name__ == '__main__':
    main()
