#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   creat_excel.py
@Time    :   2019/10/09 12:04:27
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib

import xlwings as xw
import argparse

parser = argparse.ArgumentParser(description = 'create excel in your select path')

parser.add_argument('-i','--input', metavar='', type=str, required = True, help = 'where your want to creat excel')

args = parser.parse_args()

if __name__ == '__main__':
    try:
        app = xw.App(visible=True,add_book=False)
        # wb = app.books.add()
        wb = app.books.add()
        wb.sheets["sheet1"].name = "I-V Performance"
        sht = wb.sheets['I-V Performance']
        name = [
                "file name",  
                "Cell Area", 
                "Voc(V)", 
                "Isc(A)", 
                "Vm(V)", 
                "Im(A)", 
                "Pmax(W)",
                "Efficiency(%)",
                "Fill Factor(%)", 
                "Jsc(mA/cm2)", 
                "Rs/ohm", 
                "Rsh/ohm", 
                "Light Intensity(W/m2)", 
                "Cell Temperature deg.(C)"
                ]
        sht.range('A1','N1').value = name
        # 格式化
        # 对表格进行美化
            # 对第一行标题进行格式化
        sht.range('A1').expand('right').api.HorizontalAlignment = -4108
        sht.range('A1').expand('right').api.VerticalAlignment = -4108
            # 行高
        sht.api.Rows(1).RowHeight = 20
            # 列宽
        sht.api.Columns("A:N").Columnwidth = 15
        print('表 1 创建完成')
        # 创建表 2
        wb.sheets.add("sheet2")
        wb.sheets["sheet2"].name = "refine data"
        sht2 = wb.sheets['refine data']
        name = [
                "file name",  
                "Cell Area(cm2)", 
                "Voc(V)", 
                "Jsc(mA/cm2)", 
                "Fill Factor(%)", 
                "Rs(ohm)", 
                "Rsh(ohm)", 
                ]
        sht2.range('A1','G1').value = name
        # 格式化
        # 对表格进行美化
            # 对第一行标题进行格式化
        sht2.range('A1').expand('right').api.HorizontalAlignment = -4108
        sht2.range('A1').expand('right').api.VerticalAlignment = -4108
            # 行高
        sht2.api.Rows(1).RowHeight = 20
            # 列宽
        sht2.api.Columns("A:G").Columnwidth = 15
        print('表 2 创建完成')
        # 创建表 3
        wb.sheets.add("raw data")
        # wb.sheets["sheet2"].name = "refine data"
        sht3 = wb.sheets['raw data']
        name = [
                "file name",  
                "Cell Area(cm2)", 
                "Voc(V)", 
                "Jsc(mA/cm2)", 
                "Fill Factor(%)", 
                "Rs(ohm)", 
                "Rsh(ohm)", 
                ]
        sht3.range('A1','G1').value = name
        # 格式化
        # 对表格进行美化
            # 对第一行标题进行格式化
        sht3.range('A1').expand('right').api.HorizontalAlignment = -4108
        sht3.range('A1').expand('right').api.VerticalAlignment = -4108
            # 行高
        sht3.api.Rows(1).RowHeight = 20
            # 列宽
        sht3.api.Columns("A:G").Columnwidth = 15
        print('表 3 创建完成')
    finally:
        if wb:
            wb.save(args.input)
            wb.close()
            app.kill()