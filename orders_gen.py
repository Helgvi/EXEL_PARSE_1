import sqlite3

import xlrd
import xlwt





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
