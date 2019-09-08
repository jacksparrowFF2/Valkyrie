#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   get.py
@Time    :   2019/09/02 14:29:08
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib


import xlwings as xw

global app, wb, sht, info, row, rowl
# 开始对EXCEL进行编辑
# 方式1——隐性：EXCEL在后台运行
# 创建app进程
app = xw.App(visible=False, add_book=False)
# 链接工作表,填写要写入的EXCEL文件路径
wb = app.books.open(
    'F:\github_graduate\Valkyrie\idea\\test.xlsx')
# 对指定工作表进行编辑
sht = wb.sheets['Ratio MetaData']
# 获取当前EXCEL表格的行数与列数
info = sht.range('A1').expand('table')
print(info)
row = info.last_cell.row
col = info.last_cell.column
# 计算出要添加的一行位置

rowl = str(row + 1)
print('数据添加所在行：'+rowl)
row = str(row)
print('原表格最后一行：'+row)

""" #保存文件
   wb.save()
   # 关闭文件
   wb.close()
   # 结束进程
   app.kill
    """
