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
import argparse
import win32clipboard as w
import win32con
import numpy as np
import xlwings as xw
import math

parser = argparse.ArgumentParser(description="该脚本用于对测试数据型号1数据进行整合并处理，\
    搭配Quicker进行使用效果更佳")

parser.add_argument('-i','--input data',metavar = '',type = str,
                    help = '原始数据的 Excel 表格')
parser.add_argument('-s','--save',metavar = '', type = str,
                    help = '保存到指定整合数据的 Excel 文件路径')

# # 创建基础选项
# parser.add_argument_group('基础选项')

group = parser.add_argument_group('高级选项')
group.add_argument('-W','--write', action='store_true',help='将测试数据写入到指定的 Excel 文件')
group.add_argument('-C','--Creat', action='store_true',help='创建用于整合数据的 Excel 文件路径')
args = parser.parse_args()

if __name__ == '__main__':
    if args.Creat:
        print('开始创建 Excel')
        a = args.save
        print(a)
        name = ['name','item1','item2','item3','item4']
        try:
            app = xw.App(visible=False,add_book=False)
            wb = app.books.add()
            wb.sheets['sheet1'].name = 'SUM data'
            sht = wb.sheets['SUM data']
            sht.range('A1','E1').value = name
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
    else:
        print('请选择要执行的命令')