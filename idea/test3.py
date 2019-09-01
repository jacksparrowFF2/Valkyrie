#!/usr/bin/python3
# -*- encoding: utf-8 -*-
'''
@File    :   test3.py
@Time    :   2019/09/01 10:39:13
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2017-2018, Liugroup-NLPR-CASIA
@Desc    :   None
'''

# here put the import lib

import xlwings as xw

wb = xw.Book('Report_Quartz_2019_Condition.xlsx')

# sht = wb.sheets['Raman Metadata']
sht = wb.sheets['Ratio Metadata']
# sht = wb.sheets['sheet1']
# sht = wb.sheets['Sheet2']

# a = sht.range('A1').value
# a = sht.range('A2','AF2').value
# print(a)

# 第一种方法（有缺陷）
info = sht.used_range
print(info)

# nrows = info.last_cell.row
# ncols = info.last_cell.column

# print(nrows)
# print(ncols)

# b = nrows + 1
# c = sht.range('A'+str(b),'AF'+str(b)).value

# print(c)

# #第二种获取数据表大小(有缺陷)
# info2 = sht.api.UsedRange

# nrows2 = info2.Rows.count
# ncols2 = info2.Columns.count

# print(info2)
# print(nrows2)
# print(ncols2)

#第三种方法c

info3 = sht.range('A1').expand('table')
print(info3)
# info = sht.range('A1').expand('table')
# row = info.last_cell.row
# col = info.last_cell.column
# print(info,row,col)

