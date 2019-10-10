#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   extra_IV_Y.py
@Time    :   2019/10/10 16:37:05
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import os
import xlwings as xw
import win32clipboard as w
import win32con
import argparse

parser = argparse.ArgumentParser('汇总不同条件下的性能指标')
parser.add_argument('-i','--input',metavar='',type=str,
                    help='需要汇总的 Excel 表格')
group = parser.add_argument_group('基础选项')
group.add_argument('-e','--excel', metavar='',type=str,
                   help='写入数据的 Excel 文件路径')

group = parser.add_argument_group('高级选项')
group.add_argument('-w','--write', action = 'store_true',
                   help='将汇总好的实验数据写入至指定 Excel 文件')
group.add_argument('-c','--copy', action = 'store_true',
                   help='将汇总好的实验数据写入至剪贴板')
args = parser.parse_args()

def summerize(a,b):
    sht = wb.sheets[a]
    info = sht.range('A1').expand('table')
    row = info.last_cell.row
    col = info.last_cell.column
    coll = col + 1
    if row == 1:
        sht.range(1,col).options(transpose = True).value = b
            # 调整列宽
        sht.api.columns(col).Columnwidth = 15
    else:
        sht.range(1,coll).options(transpose = True).value = b
            # 调整列宽
        sht.api.columns(coll).Columnwidth = 15
    # 格式化
        # 更新行数
    info = sht.range('A1').expand('table')
    row = info.last_cell.row
        # 居中对齐
    sht.range('A1').expand('table').api.HorizontalAlignment = -4108
    sht.range('A1').expand('table').api.VerticalAlignment = -4108
        # 调整行高
    sht.range('A1:A'+str(row)).api.Rowheight = 20
    

if __name__ == '__main__':
    if args.write:
        # 读取数据
        inexcel = args.input
        try:
            app = xw.App(visible = False, add_book = False)
            wb = app.books.open(inexcel)
            sht = wb.sheets['refine data']
            info = sht.range('A1').expand('table')
            row = info.last_cell.row            
            col = info.last_cell.column
            
            item = os.path.split(str(inexcel))[1].split('.')[0]
            print(item)
            
            PCE = sht.range('B2').options(expand='down').value
            Voc = sht.range('C2').options(expand='down').value
            Jsc = sht.range('D2').options(expand='down').value
            FF = sht.range('E2').options(expand='down').value
            Rs = sht.range('F2').options(expand='down').value
            Rsh = sht.range('G2').options(expand='down').value 
            
            PCE.insert(0,item)
            Voc.insert(0,item)
            Jsc.insert(0,item)
            FF.insert(0,item)
            Rs.insert(0,item)
            Rsh.insert(0,item)
            
            print(Voc)
            print(Jsc)
            print(FF)
            print(Rs)
            print(Rsh)
        finally:
            if wb:
                # 保存文件
                wb.save()
                # 关闭文件
                wb.close()
                # 结束进程
                app.kill()
        # 写入数据
        outexcel = args.excel
        try:
            app = xw.App(visible = False, add_book = False)
            wb = app.books.open(outexcel)
            
            summerize("PCE",PCE)
            summerize("Voc",Voc)
            summerize("Jsc",Jsc)
            summerize("FF",FF)
            summerize("Rs",Rs)
            summerize("Rsh",Rsh) 
        finally:
            if wb:
                # 保存文件
                wb.save()
                # 关闭文件
                wb.close()
                # 结束进程
                app.kill()
    elif args.copy:
        print(0)
    else:
        print("请输入 -h 以查看帮助")
        input("Press <enter>")