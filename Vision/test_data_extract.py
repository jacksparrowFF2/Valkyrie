#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   test_data_extract.py
@Time    :   2021/07/31 20:34:26
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2021, EXphysiclab
@Desc    :   None
'''
# 创建header头文件

# 导入依赖包
# here put the import lib
import os
import math
import win32con
import argparse
import win32clipboard as w
import numpy as np
import xlwings as xw

from xlwings.main import App, Sheet, Sheets

parser = argparse.ArgumentParser(description="该脚本用于对测试数据型号1数据进行整合并处理，\
    搭配Quicker进行使用效果更佳")

parser.add_argument('-i', '--input', metavar='', type=str,
                    help='原始数据的 Excel 表格')
parser.add_argument('-s', '--save', metavar='', type=str,
                    help='保存到指定整合数据的 Excel 文件路径')

# # 创建基础选项
# parser.add_argument_group('基础选项')

group = parser.add_argument_group('高级选项')
group.add_argument('-W', '--Write', action='store_true',
                    help='将测试数据写入到指定的 Excel 文件')
group.add_argument('-C', '--Creat', action='store_true',
                    help='创建用于整合数据的 Excel 文件路径')
args = parser.parse_args()

if __name__ == '__main__':
    if args.Creat:
        print('开始创建 Excel')
        a = args.save
        print(a)
        name = ['name', 'item1', 'item2', 'item3', 'item4']
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.add()
            wb.sheets['sheet1'].name = 'Sum Data'
            sht = wb.sheets['Sum Data']
            sht.range('A1', 'E1').value = name
            # 格式化
            # 对表格进行美化
            # 对第一行标题进行格式化
            sht.range('A1').expand('right').api.HorizontalAlignment = -4108
            sht.range('A1').expand('right').api.VerticalAlignment = -4108
            print('格式化完成')
        finally:
            if wb:
                wb.save(args.save)
                wb.close()
                app.kill()
    elif args.Write:
        # 整合文件路径赋值给 infile
        inexcel = args.save
        print('输入的整合文件路径为：'+inexcel)
        # 输入的数据文件路径
        indata = args.input
        print('输入的数据文件路径为：'+indata)
        # 数据中转区
        test_data = []
        # 屏体ID
        ID = os.path.split(str(indata))[1].split('.')[0]
        print(ID)
        print(type(ID))
        # 数据矩阵大小
        # 数据矩阵行数，初始化为 0 
        data_row = 0
        # 数据矩阵列数，初始化为 0
        data_col = 0
        # 打开数据文件
        try:
            app = xw.App(visible=True, add_book=False)
            wb = app.books.open(indata)
            sht = wb.sheets['Sheet1']
            # 获取表格坐标信息
            info = sht.range('A1').expand('table')
            row = info.last_cell.row
            col = info.last_cell.column
            print('原表格最后一行：'+str(row))
            print('原表格最后一列：'+str(col))
            data_row = row
            data_col = col
            # 提取某一单元格的值
            a = sht.range((2, 1)).value
            print(a)
            # 提取[XY,MP]范围内单元格的值
            b = sht.range((2, 1), (row, col))
            print(b)
            test_data = b.value
            print(test_data)
            print('数据复制完毕')
        except:
            if wb:
                wb.save()
                wb.close()
                app.kill()
            print('程序异常')
        finally:
            if wb:
                wb.save()
                wb.close()
                app.kill()
            print('数据读取完毕')
        print(test_data)
        # 打开汇合总表并写入数据中转区列表
        try:
            app = xw.App(visible=True,add_book=False)
            wb = app.books.open(inexcel)
            sht = wb.sheets['Sum Data']
            # 获取表格坐标信息
            # info = sht.range('A1').expand('table')
            # row = info.last_cell.row
            # col = info.last_cell.column
            # print('汇总表格最后一行：'+str(row))
            # print('汇总表格最后一列：'+str(col))
            
            info = sht.used_range
            row = info.last_cell.row
            col = info.last_cell.column
            print('汇总表格实际最后一行：'+str(row))
            print('汇总表格实际最后一列：'+str(col))

            # 填写屏体编号
            sht.range((row+1,1)).value = ID
            # 将中转区列表写入同大小单元格区域
            # 将[XY,MP]范围内单元格的值复制到另一个区域
            sht.range((row+1,2)).options(expand = 'table').value = test_data
            print('数据粘贴完毕')
            # 合并单元格
            sht.range((row+1,1),(row+data_row-1,1)).merge()
            
            # # 填写屏体编号
            # sht.range((row+1,1)).value = ID
            # # sht.range((row+1,1),(row+data_row-1,1)).options(transpose = True).value = [ID]*6
        except:
            if wb:
                wb.save()
                wb.close()
                app.kill()
            print('写入数据异常')
        finally:
            if wb:
                wb.save()
                wb.close()
                app.kill()
            print('数据写入完毕')
    else:
        print('请选择要执行的命令')
