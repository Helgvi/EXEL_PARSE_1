from xlwt import *

w = Workbook()
ws = w.add_sheet('Hey, Dude')

for i in range(6, 80):
    fnt = Font()
    fnt.height = i*20
    style = XFStyle()
    style.font = fnt
    ws.write(1, i, 'Test')
    ws.col(i).width = 0x0d00 + i
w.save('col_width.xls')
