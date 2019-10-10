#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   createxcel.py
@Time    :   2019/09/22 19:20:25
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
        app = xw.App(visible=False,add_book=False)
        # wb = app.books.add()
        wb = app.books.add()
        # wb.sheets.add("sheet2")
        wb.sheets["sheet1"].name = "I-V"
        # xw.sheets.add(name = 'I-V')
        sht = wb.sheets['I-V']
        name = ["Code", "NO.", "Time/s", "Serial NO.", "Voc/V", "Isc/mA", "Pmax/mW", "Vpmax/V", "Ipmax/mA", "Rs/ohm", "Rsh/ohm", "Jsc/mA.cm-2", "FF", "η/%"]
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
        print('格式化完成')
    finally:
        if wb:
            wb.save(args.input)
            wb.close()
            app.kill()