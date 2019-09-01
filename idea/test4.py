#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   test4.py
@Time    :   2019/09/01 14:03:41
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2017-2018, Liugroup-NLPR-CASIA
@Desc    :   None
'''

# here put the import lib
import xlwings as xw
#1


#2
app = xw.App(visible=False,add_book=False)
wb = app.books.open('Report_Quartz_2019_Condition.xlsx')
sht = wb.sheets['Ratio Metadata']
#获取当前EXCEl表格的行数与列数
info = sht.range('A1').expand('table')
print(info)
row = info.last_cell.row
col = info.last_cell.column
print(row)
print(col)
# #输出最下面一行的数据
a = row + 1
b = str(a)

c = sht.range('A'+b,'AF'+b).value
# print(c)
# #在最下面一行指定位置添加新数据
d = [1,2,3,7]
sht.range('N'+b,'V'+b).value = d
# # #输出更改后矩阵的大小
# wb.save()
# info2 = sht.range('A1').expand('table')
# row2 = info.last_cell.row
# col2 = info.last_cell.column
# print(row2)
# print(col2)
# #保存数据并关闭
# wb.save()
wb.close()
app.kill()

