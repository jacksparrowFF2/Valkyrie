#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   extract_IV.py
@Time    :   2019/09/21 21:13:08
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
# 引入命令行组件
import argparse
import xlwings
import sys
import win32clipboard as w
import win32con

parser = argparse.ArgumentParser(description='This script is aims to extract I-V data from text file')

# parser.add_argument('-i','--input',metavar='',type=argparse.FileType(mode='r'),
#                     help='the file you want to extract')
parser.add_argument('-i', '--input', metavar = '', type=argparse.FileType(mode='r'),
                    help='the file you want to extract')
parser.add_argument('-e', '--excel', metavar = '', type= str,
                    help='the file you want to extract')

group = parser.add_argument_group('Basic Options')
group.add_argument('-t','--time', action = 'store_true',
                   help='the time of test')
group.add_argument('-a','--area', action = 'store_true',
                   help='the area of device')
group.add_argument('-m','--material', action = 'store_true',
                   help='the type of device')

group = parser.add_argument_group('Advanced Options')
group.add_argument('-w','--write', action = 'store_true',
                   help='write to excel you select')
group.add_argument('-c','--copy', action = 'store_true',
                   help='copy data to your clip')

args = parser.parse_args()

def writeclip(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, astring)
    w.CloseClipboard()

def str2list(astring):
    in_str = astring.replace('\n', '')
    temp_list = in_str.split('\t')
    out_list = []
    for i in temp_list:
        i = i.lstrip()
        out_list.append(i)
    return(out_list)

if __name__ == '__main__':
    # print('请输入 -h 以查看使用帮助')
    if args.copy:
        infile = args.input
        All_data = infile.readlines()
        print(All_data[10])
        writeclip(All_data[10])
    elif args.write:
        # name = str2list(All_data[9])
        # print(name)
        data = str2list(All_data[10])
        print(data)
    else:
        print('请输入 -h 以查看帮助')
    
    # # print(All_data)
    # time = All_data[2]
    
    # name = str2list(All_data[9])
    # print(name)
    
    # data = str2list(All_data[10])
    # print(data)

