import sqlite3
import xlrd
import xlwt


ANSWER = 'Код 1С Не найден!'
TITLE_LIST = [
    'Идентификатор',
    'Код в 1C',
    'Артикул произв',
    'Наименование',
    'Количество',
    'Цена',
    'Сумма'
    ]


path_intake = 'c:/T000582215.xls'
path2_exhtake = 'c:/New_list.xls'

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


def answer_from_exel_file():
    rb = xlrd.open_workbook(path_intake)
    print("Листов книги Exel - {0}".format(rb.nsheets))
    print("Листы файла: {0}".format(rb.sheet_names()))
    sheet = rb.sheet_by_index(0)
    num = sheet.nrows
    title = sheet.cell(0, 0).value
    totals = sheet.cell(num-1, 9).value
    for rx in range(4, num-1):
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
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Накладная")
    sheet1.write(0, 0, titles_list[0])
    for num in range(7):
        sheet1.write(1, num, TITLE_LIST[num])
    for num in range(2, len(code_list)):
        sheet1.write(num, 0, code_list[num])
        sheet1.write(num, 1, after_action_list[num])
        sheet1.write(num, 2, art_list[num])
        sheet1.write(num, 3, name_list[num])
        sheet1.write(num, 4, quent_list[num])
        sheet1.write(num, 5, price_list[num])
        sheet1.write(num, 6, amount_list[num])
    sheet1.write(len(code_list)+1, 6, titles_list[1])
    book.save(path2_exhtake)


def main():
    answer_from_exel_file()
    bild_list()
    return_1C_code(code_list)
    print("Число записей - {}".format(len(code_list)))
    print("Число записей - {}".format(len(after_action_list)))
    write_new_data()


if __name__ == '__main__':
    main()
