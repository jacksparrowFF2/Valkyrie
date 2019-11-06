#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   extract_XRD_CZTS.py
@Time    :   2019/11/06 21:23:06
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
from operator import itemgetter
import win32clipboard as w
import win32con
import xlwings as xlwings
import argparse

# 创建命令解释器
parser = argparse.ArgumentParser('该脚本旨在帮助你更方便的对 XRD 衍射数据进行整理')

parser.add_argument('-i','--input', metavar = '', type = argparse.FileType(mode='r'), 
                    help = '要进行整理的 XRD 数据文件路径')
parser.add_argument('-e','--excel', metavar = '', type = str, 
                    help = '保存数据的 excel 文件路径')

group = parser.add_argument_group('基础选项')
group.add_argument('-c','--column', metavar = '', type = int, 
                   help = '要提取的数据列')

group = parser.add_argument_group('进阶选项')
group.add_argument('-wc','--wcondition', action = 'store_true', 
                   help = '模式：将 XRD 结果写入指定的 excel 表格')
group.add_argument('-wx','--wxrd', action = 'store_true', 
                   help = '模式：将 XRD 结果写入指定的 excel 表格')
group.add_argument('-cs','--copyselect', action = 'store_true', 
                   help = '模式：将提取的 XRD 数据列写入剪贴板')

args = parser.parse_args()

if __name__ == '__main__':
    if args.wcondition:
        infile = args.input
        All_data = infile.readlines()
        in_xrddate = All_data[140:]
        print(in_xrddate)
    elif args.wxrd:
        print('2')
    elif args.copyselect:
        print('3')
    else:
        print('请输入 -h 以查看使用说明')
        input("Press <enter>")
    