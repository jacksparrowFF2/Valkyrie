#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   test6.py
@Time    :   2019/09/17 20:00:48
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
# a = [1,2,3]
# b = a[1:2]
# print(b)

import xlwings as xw
path = r'F:\github_graduate\Valkyrie\idea\test.xlsx'
print(path)
# print(type(path))

# 开始对EXCEL进行编辑
# 方式1——隐性：EXCEL在后台运行
# 创建app进程
app = xw.App(visible=False, add_book=False)
# 链接工作表,填写要写入的EXCEL文件路径
try:
    wb = app.books.open(path)
    # 对指定工作表进行编辑
    sht = wb.sheets['Raman MetaData']
    # 获取当前EXCEL表格的行数与列数
    info = sht.range('A1').expand('table')
    # print(info)
    row = info.last_cell.row
    col = info.last_cell.column
    # 计算出要添加的一行位置

    row = info.last_cell.row
    col = info.last_cell.column
    # 计算出要添加的一行位置
    coll = col + 1
    str_coll = str(coll)
    print('数据添加所在列：'+str_coll)
    str_col = str(col)
    print('数据添加所在列：'+str_col)
    # 注入数据
    # sht.range('C2').value = 1
    # sht.cells(3).value = 3
    # 填充序列
    a = [[1],[2],[3],[4],[5]]
    # print(a)
    # print(type(a))
    # print(type(a[0]))
    
    # b = list(range(5,10))
    
    # a = '=INDIRECT('
    # print(a)
    # b = '"'
    # print(b)
    # c = '\'Ratio Metadata\''
    # print(c)
    # d = '!$A'
    # print(d)
    # e = '"&COLUMN())'
    # print(e)

    # formula = a + b + c + d + e
    # print(formula)
    
    # b  = ['136', '144.006', '136.011', '147.017', '144.023']
    # b.insert(0,formula)
    # print(b)
    
    # 单个单元格赋值cellls(row,col)
    # sht.cells(1,3).value = 100
    # sht.cells(5,3).value = 100
    
    # 多个单元格赋值 range(row,col)
    #     # 从第 3 列 第 1 行 开始往下赋值，a 有多少个元素，就赋值到多少行
    # sht.range(1,3).value = a
    #     # 从第 4 列 第 1 行 开始往下赋值至 第 4 列，第 5 行，a的元素应与行数相同
    # sht.range((1,4),(5,4)).value = a
        # 从第 5 列 第 1 行 开始往下赋值至 第 5 列，第 5 行，a的元素应与行数相同
    # sht.range((1,coll),(row,coll)).options(transpose = True).value = b
    
    # 单元格格式化
        # 对指定单元格进行水平+垂直居中对齐
    # sht.range((1,1).api.HorizontalAlignment= -4108
    # sht.range((1,1).api.VerticalAlignment= -4108
        # 对指定范围进行水平+垂直居中对齐 坐标形式(行，列)
    # sht.range((1,2),(3,2)).api.HorizontalAlignment= -4108
    # sht.range((1,2),(3,2)).api.VerticalAlignment= -4108
        # 对指定列进行水平+垂直居中，有两种方式，此为方式 1
    # sht.api.columns(1).HorizontalAlignment= -4108
    # sht.api.columns(1).VerticalAlignment= -4108
        # 对指定列进行水平+垂直居中，有两种方式，此为方式 2
    # sht.range('A1').expand('down').api.VerticalAlignment = -4108
    # sht.range('A1').expand('down').api.HorizontalAlignment = -4108
    # sht.range('B2:C2').expand('down').api.HorizontalAlignment = -4108
    # sht.range('A1').api.style = "Percent"
    # sht.range('A2').api.NumberFormat = "##.00_)"
    # sht.range('A3').api.ShrinkToFit = True
    # sht.range('A4').api.WrapText = True
    # sht.range('A5').api.NumberFormat = "000"
    # sht.range('A6').api.NumberFormat = "00"
    sht.api.columns(4).ColumnWidth = 40
    
    
    # 写入公式
    # sht.range((1,4)).formula = r'ss1!'=INDIRECT("'Ratio Metadata'!$A"&COLUMN())
    
finally:
    if wb:
        #保存文件
        wb.save()
        # 关闭文件
        wb.close()
        # 结束进程
        app.kill




































# wb = app.books.open(path)
# # 对指定工作表进行编辑
# sht = wb.sheets['Raman MetaData']
# # 获取当前EXCEL表格的行数与列数
# info = sht.range('A1').expand('table')
# print(info)
# row = info.last_cell.row
# col = info.last_cell.column
# # 计算出要添加的一行位置

# row = info.last_cell.row
# col = info.last_cell.column
# # 计算出要添加的一行位置
# coll = col + 1
# str_coll = str(coll)
# print('数据添加所在列：'+str_coll)
# str_col = str(col)
# print('数据添加所在列：'+str_col)
# # 设置单元格格式
# # sht.cells
# # 注入数据
# # sht.range('C2').value = 1
# # sht.cells(3).value = 3
# # 填充序列
# # sht.range('C1').value = [[1],[2],[3],[4],[5]]


# #保存文件
# wb.save()
# # 关闭文件
# wb.close()
# # 结束进程
# app.kill