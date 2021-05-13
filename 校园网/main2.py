import xlwt
# workbook = xlwt.open_workbook(data.xls)
# workbook.write(0, 0, label = 'Row 0, Column 0 Value')
# print( "第一行第一列的值为:", sh1.cell_value(0, 0))


write_hang = 0
# for a in range(5):
while write_hang < 10:
    write_wb = xlwt.Workbook()
    write_sh1 = write_wb.add_sheet('miyao')
    write_sh1.write(write_hang, 0, '1')
    write_sh1.write(write_hang, 1, '1')
    write_hang = write_hang + 1
    print(write_hang)
write_wb.save('./data.xls')