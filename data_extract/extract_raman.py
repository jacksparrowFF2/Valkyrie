#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   extract_raman.py
@Time    :   2019/09/17 15:14:56
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
# 导入剪贴板组件
import win32clipboard as w
import win32con
# 导入命令行参数组件
import argparse
# 导入excel操作组件
import xlwings as xw
# 创建命令解释器
parser = argparse.ArgumentParser(description='This script is aims to extract Raman date from txt file')
# 创建命令行输入参数，输入参数为文件路径
# parser.add_argument("-i","--input", type=argparse.FileType(mode = 'r'), required = True, 
#                      help = 'the file need to extract data')
# parser.add_argument("-e","--excel", type=str, required = True, 
#                      help = 'the file need to extract data')
parser.add_argument("-i","--input", type=argparse.FileType(mode = 'r'),
                     help = 'the file need to extract data')
parser.add_argument("-e","--excel", type = str, help = 'the file need to open')
# 创建附属命令行参数，增加可选输出第二列的选项
group = parser.add_argument_group('Basic options')
group.add_argument('-c','--column', type = int, help = 'chose the column you want to extract')
# 创建互斥锁
group = parser.add_mutually_exclusive_group()
# group = parser.add_argument_group('advanced options')
group.add_argument('-a','--all', action = 'store_true', help = 'this will extract all data to your clipboard')
group.add_argument('-s','--select', action = 'store_true', help = 'this will only extract the select column to your clipboard')
group.add_argument('-r','--write', action = 'store_true', help = 'this will add your raman data to your excel file last column')

args = parser.parse_args()


# 创建剪贴板写入函数
def writeclip(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, astring)
    w.CloseClipboard()
    
    
if __name__ == '__main__':
    # infile = args.input
    # filecontents = infile.read()
    if args.write:
        inexcel = args.excel
        # inexcel = 'r\'%s\'' %(args.excel)
        print('你输入的文件路径为：'+inexcel)
        # # 创建 app 进程
        # app = xw.App(visible = True, add_book = False)
        # 创建 app 进程
        app = xw.App(visible = False, add_book = False)
        # 链接工作表,填写要写入的EXCEL文件路径
        wb = app.books.open(inexcel)
        # 对指定工作表进行编辑
        sht = wb.sheets['Raman MetaData']
        # 获取当前EXCEL表格的行数与列数
        info = sht.range('A1').expand('table')
        print(info)
        row = info.last_cell.row
        col = info.last_cell.column
        # 计算出要添加的一行位置
        coll = str(col + 1)
        print('数据添加所在列：'+coll)
        col = str(col)
        print('数据添加所在列：'+col)
        # 注入实验数据
        # sht.range('O'+rowl, 'X'+rowl).value = data
        #保存文件
        wb.save()
        # 关闭文件
        wb.close()
        # 结束进程
        app.kill
    elif args.all:
        infile = args.input
        filecontents = infile.read()
        print("this is all experiment data you get from test, you can find it in your clipborad")
        print(filecontents)
        writeclip(filecontents)
    elif args.select:
        infile = args.input
        filecontents = infile.read()
        print("you select column is %s" %(args.column))
        # 转化至程序排序方式
        n = args.column - 1
        # 将字符串转换为列表，以换行符为切割处
        select_list = filecontents.split('\n')
        # 调试输出
        # print(select_list)
        # 构建格式化列表
        format_select_list = []
        for i in select_list:
            format_select_list.append(i.split('\t'))
        # 调试输出
        # print(format_select_list)
        # print(len(format_select_list))
        # 构建输出列表
        out_select_list =[]
        for i in range(len(format_select_list)):
            # print(format_select_list[i][n])
            out_select_list.append(format_select_list[i][n])
            # out_select_list[i] = format_select_list[i][n]
        print(out_select_list)
        # 构建输出字符串
        str_data = "\n".join(out_select_list)
        # 调试输出
        print(str_data)
        writeclip(str_data)
    else:
        # infile = args.input
        # filecontents = infile.read()
        # # 调试输出
        # print(filecontents)
        # writeclip(filecontents)
        print(5)


    # print(filecontents)
    # print(type(filecontents))
    # writeclip(filecontents)
    