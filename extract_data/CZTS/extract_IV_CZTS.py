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
# 引入Excel操作组件
import xlwings as xw
# 引入剪贴板组件
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
        print(All_data[10:])
        writeclip(All_data[10])
    elif args.write:
        infile = args.input
        All_data = infile.readlines()
        # name = str2list(All_data[9])
        # print(name)
        data = str2list(All_data[10])
        print(data)
        """ try:
            inexcel = args.excel
            app = xw.App(visible=False,add_book = False)
            wb = app.books.open(inexcel)
            sht = wb.sheets['I-V']
            info = sht.range('A1').expand('table')
            row = info.last_cell.row
            col = info.last_cell.column
            rowl = row + 1
            print('原表格最后一行：'+str(row))
            print('数据添加所在行：'+str(rowl))
            # 注入数据
            sht.range('A'+str(rowl)).value = 'TAG'
            sht.range('B'+str(rowl),'N'+str(rowl)).value = data
            print('注入完成')
            # 格式化
                # A列进行自动换行+粗体+右对齐+垂直居中
            sht.range('A'+str(rowl)).api.WrapText = True
            sht.range('A2').expand('down').api.font.Bold = True
            sht.range('A2').expand('down').api.HorizontalAlignment = -4152
            sht.range('A2').expand('down').api.VerticalAlignment = -4108
                # B:N 列进行垂直水平居中对齐
            sht.range('B'+str(rowl),'N'+str(rowl)).api.HorizontalAlignment = -4108
            sht.range('B'+str(rowl),'N'+str(rowl)).api.VerticalAlignment = -4108
                # 行高
            sht.api.Rows(rowl).RowHeight = 20
            print('格式化完成')
            
        finally:
            if wb:
                wb.save()
                wb.close()
                app.kill() """
    else:
        print('请输入 -h 以查看帮助')
    
    # # print(All_data)
    # time = All_data[2]
    

    # name = str2list(All_data[9])
    # print(name)
    
    # data = str2list(All_data[10])
    # print(data)

