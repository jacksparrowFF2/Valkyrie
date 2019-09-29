#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   format_raman_result.py
@Time    :   2019/09/29 09:16:17
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import xlwings as xw
import win32clipboard as w
import win32con
import argparse
from operator import itemgetter

parser = argparse.ArgumentParser(description = '将拟合好的拉曼结果按照规定的方式填写入Excel表格中')
parser.add_argument('-i','--input', metavar = '', type = str, help = '要写入的excel文件夹路径')
group = parser.add_argument_group('基本选项')
group.add_argument('-s','--select',action = 'store_true', help = '输出指定数据')
group.add_argument('-a','--all',action = 'store_true', help = '输出所有数据')
args = parser.parse_args()

def getclip():
    w.OpenClipboard()
    copy_text = w.GetClipboardData(win32con.CF_UNICODETEXT)
    w.CloseClipboard()
    return copy_text

def writeclip(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, astring)
    w.CloseClipboard()
def rformat(astring):
    a = astring.split('\n')
    while '' in a:
        a.remove('')
    c = []
    for item in a:
        b = item.split('\t')
        c.append(b)
    d = sorted(c,key=itemgetter(5))
    return(d)



if __name__ == '__main__':
    if args.select:
        temp_list = rformat(getclip())
        out_list = []
        for i in temp_list:
            out_list.append(i[2]+'\t')
            out_list.append(i[3]+'\t')
        print(out_list)
        del out_list[8:10]
        # del out_list[-1]
        out_str = "".join(out_list)
        print(out_str)
        writeclip(out_str)
    elif args.all:
        temp_list = rformat(getclip())
        for i in temp_list:
            i[6] = i[6]+'\n'
        print(temp_list)
        out_list = []
        for i in temp_list:
            out_list.append("\t".join(i))
        print(out_list)
        out_str = "".join(out_list)
        print(out_str)
    else:
        print('test')