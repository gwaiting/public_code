from time import monotonic
import xlrd
import xlwt





# 打开刚才我们写入的 test_w.xls 文件
wb = xlrd.open_workbook("./data.xls")
# 获取并打印 sheet 数量
print( "sheet 数量:", wb.nsheets)
# # 获取并打印 sheet 名称
# print( "sheet 名称:", wb.sheet_names())
# 根据 sheet 索引获取内容
list = []
a = 0
while a < 400:
    sheet = wb.sheet_by_index(0)
    print(sheet)
    name = sheet.cell_value(a, 0)
    code = sheet.cell_value(a, 1)
    mobil_url = 'http://10.200.84.3:801/eportal/?c=Portal&a=login&callback=dr1003&login_method=1&user_account={}&user_password={}&wlan_user_ip=你的IP地址&wlan_user_ipv6=&wlan_user_mac=000000000000&wlan_ac_ip=&wlan_ac_name=&jsVersion=3.3.2&v=7679'.format(name, code)
    pc_url = 'http://10.200.84.3:801/eportal/?c=Portal&a=login&callback=dr1003&login_method=1&user_account={}&user_password={}&wlan_user_ip=10.190.86.212&wlan_user_ipv6=&wlan_user_mac=000000000000&wlan_ac_ip=&wlan_ac_name=&jsVersion=3.3.2&v=8972'.format(name, code)
    # print(name)
    # print(code)
    # print(mobil_url)
    list.append(pc_url)
    row = 0
    book = xlwt.Workbook()
    sheet = book.add_sheet('miyao')
    for name in list:
        sheet.write(row, 0, name)
        row = row + 1
    book.save('./url_pc.xls')
    print('账号已写入')
    print(pc_url)
    a = a + 1