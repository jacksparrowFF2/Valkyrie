#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   Untitled-1
@Time    :   2019/09/02 11:09:08
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, Liugroup-NLPR-CASIA
@Desc    :   None
'''

# here put the import lib
import xlwings as xw
global app,wb,sht,info,row,rowl
#开始对EXCEL进行编辑
#创建app进程
app = xw.App(visible=False,add_book = False)
# 链接工作表
wb = app.books.open('Report_Quartz_2019_Condition.xlsx')
# 对指定工作表进行编辑
sht = wb.sheets['Ratio MetaData']
# 方式2——显性
# wb = xw.Book('Report_Quartz_2019_Condition.xlsx')
# sht = wb.sheets['Ratio Metadata']
# 获取当前EXCEL表格的行数与列数
info = sht.range('A1').expand('table')
print(info)
row = info.last_cell.row
col = info.last_cell.column
# 计算出要添加的一行位置

rowl =str(row + 1)
print('数据添加所在行：'+rowl)
row = str(row)
print('原表格最后一行：'+row)
