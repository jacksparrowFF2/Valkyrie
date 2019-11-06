#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   extra_creat_excel_Y.py
@Time    :   2019/10/10 16:40:55
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import xlwings as xw
import argparse
import os
import time

parser  = argparse.ArgumentParser(description = '创建 2 阶汇总表格')

parser.add_argument('-i','--input',metavar='',type=str,required=True,
                    help='在指定位置创建EXCEL文件')
args = parser.parse_args()

# 获取当前路径并规定excel文件路径
# abspath = os.path.abspath('.')
# print(abspath)
# timetick = time.strftime("%Y_%m_%d-%H_%M_%S", time.localtime())
# filepath = abspath+'\\'+'test'+'.xlsx'
# print(filepath)

filepath = args.input
print(filepath)

try:
    app = xw.App(visible=False,add_book=False)
    wb = app.books.add()
    wb.sheets["sheet1"].name = "Rsh"
    wb.sheets.add("Rs")
    wb.sheets.add("FF")
    wb.sheets.add("Jsc")
    wb.sheets.add("Voc")
    wb.sheets.add("PCE")
    print('表格创建完成')
finally:
    if wb:
        wb.save(filepath)
        wb.close()
        wb.kill()