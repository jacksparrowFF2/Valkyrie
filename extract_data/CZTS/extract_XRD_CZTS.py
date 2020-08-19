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
import xlwings as xw
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
                   help = '模式：将 XRD 样品条件写入指定的 excel 表格')
group.add_argument('-wx','--wxrd', action = 'store_true', 
                   help = '模式：将 XRD 结果写入指定的 excel 表格')
group.add_argument('-cs','--copyselect', action = 'store_true', 
                   help = '模式：将提取的 XRD 数据列写入剪贴板')
group.add_argument('-ca','--copyall', action = 'store_true', 
                   help = '模式：将所有的 XRD 数据列写入剪贴板')

args = parser.parse_args()

# 剪贴板写入函数
def writeclip(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, astring)
    w.CloseClipboard()

# 用于去除 XRD 数据中的逗号分隔符
def remove_comma(alist):
    temp1 = []
    n = 1
    for i in alist:
        i = i.split(",")
        temp2 = []
        for element in i:
            element = element.replace(" ", "")
            temp2.append(element)
        temp1.append(temp2)
    return(temp1)

# 输出所有结果
def output_all(alist):
    temp1 = []
    for i in alist:
        i = " ".join(i)
        temp1.append(i)
    astring = "".join(temp1)
    return(astring)

# 输出指定列
def output_select(alist):
    temp1 = []
    n = args.column - 1
    if n == 1:
        n = 0
    else:
        n = 1
    for i in alist:
        i.pop(n)
        astr = "".join(i)
        temp1.append(astr)
    astr = "".join(temp1)
    return(astr)

if __name__ == '__main__':
    if args.wcondition:
        print('1')
    elif args.wxrd:
        print('2')
    elif args.copyselect:
        print('Copyselect')
        infile = args.input
        All_data = infile.readlines()
        in_xrddata = All_data[136:]
        print("你选择输出的数据列为：%s" %(args.column))
        # print(in_xrddata)
        temp = remove_comma(in_xrddata)
        output_xrddata = output_select(temp)
        print(output_xrddata)
        writeclip(output_xrddata)  
    elif args.copyall:
        print('Copyall')
        infile = args.input
        All_data = infile.readlines()
        in_xrddata = All_data[136:]
        # remove_comma(in_xrddata)
        temp = remove_comma(in_xrddata)
        output_xrddata = output_all(temp)
        writeclip(output_xrddata)
        print(output_xrddata) 
    else:
        print('请输入 -h 以查看使用说明')
        input("Press <enter>")
    