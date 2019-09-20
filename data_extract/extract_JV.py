#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   extract_JV.py
@Time    :   2019/09/17 17:13:16
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
# 创建命令解释器
parser = argparse.ArgumentParser(
    description='This script is aims to extract J-V data from txt file')
# 创建命令行输入参数，输入参数为文件路径
parser.add_argument('-i', '--input', metavar = '', type=argparse.FileType(mode='r'), required=True,
                    help='the file need to extract data')
# 创建附属命令行参数，增加可选输出第二列的选项
group = parser.add_argument_group(description = 'Basic options')
group.add_argument('-c', "--column", metavar = '', type=int,
                    help='chose the column you want to process')
# 创建互斥锁
group = parser.add_argument_group('advanced options')
# group = parser.add_mutually_exclusive_group(description = 'Basic options')
group.add_argument('-s', '--select', action='store_true',
                   help='this will only extract the select column to your clipboard')
group.add_argument('-d', '--delete', action='store_true',
                   help='this will delete you do not want data and copy the remaining data to your clipboard')
group.add_argument('-t', '--date', action='store_true',
                   help='to know when this txt file is created')
group.add_argument('-p', '--performance', action='store_true',
                   help='the performance of this solar cell')
# group.add_argument('-t', '--time', action='store_ture', 
#                    help='to know when this txt file is created')
# group.add_argument("-p", "--performance", action='store_true', 
#                    help='the performance of this solar cell')
args = parser.parse_args()


# 创建剪贴板写入函数
def writeclip(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, astring)
    w.CloseClipboard()


if __name__ == '__main__':
    # 文件路径赋值给 infile
    infile = args.input
    # 从第 30 行处开始读取 txt 文件
    All_data = infile.readlines()
    filecontents = All_data[30:]
      
    if args.select:
        print("you select column is %s" % (args.column))
        # 转化至程序排序方式
        n = args.column - 1
        # 构建格式化列表
        # # 调试输出
        # print(filecontents)
        format_select_list = []
        for i in filecontents:
            i = i.replace('\n', '')
            format_select_list.append(i.split('\t'))
        # # 调试输出
        # print(format_select_list)
        # print(len(format_select_list))
        # 构建输出列表
        out_select_list = []
        for i in range(len(format_select_list)):
            # print(format_select_list[i][n])
            out_select_list.append(format_select_list[i][n])
            # out_select_list[i] = format_select_list[i][n]
        # print(out_select_list)
        # 构建输出字符串
        str_data = "\n".join(out_select_list)
        print(str_data)
        writeclip(str_data)

    elif args.delete:
        print("the remaining data as follow show:")
        # 转化至程序排序方式
        n = args.column - 1
        remaining_list = filecontents
        if n!=2:
            print("test")
             # 构建格式化列表
            format_remaining_list = []
            for i in remaining_list:
                format_remaining_list.append(i.split('\t'))
            # # 调试输出
            # print(format_remaining_list)
            out_remaining_list = []
            # 对列表进行循环，len()函数可以获取列表的长度，列表的长度就等于数据的行数
            for i in range(len(format_remaining_list)):
                # 去除掉指定行元素
                format_remaining_list[i].pop(n)
                # 将嵌套列表转换成字符串
                temp = "\t".join(format_remaining_list[i])
                # 将转好的字符串按顺序依次添加至准备好的列表中
                out_remaining_list.append(temp)
            # 调试输出
            print(out_remaining_list)
            # 构建输出字符列表，将输出列表转换为字符串
            str_remaining_data = "".join(out_remaining_list)
            print(str_remaining_data)
            writeclip(str_remaining_data)
        else:
            # 构建格式化列表
            format_remaining_list = []
            for i in remaining_list:
                format_remaining_list.append(i.split('\t'))
            # # 调试输出
            # print(format_remaining_list)
            out_remaining_list = []
            # 对列表进行循环，len()函数可以获取列表的长度，列表的长度就等于数据的行数
            for i in range(len(format_remaining_list)):
                # 去除掉指定行元素
                format_remaining_list[i].pop(n)
                # 将嵌套列表转换成字符串
                temp = "\t".join(format_remaining_list[i])
                # 将转好的字符串按顺序依次添加至准备好的列表中
                out_remaining_list.append(temp)
            # # 调试输出
            # print(out_remaining_list)
            # 构建输出字符列表，将输出列表转换为字符串
            str_remaining_data = "\n".join(out_remaining_list)
            print(str_remaining_data)
            writeclip(str_remaining_data)
    elif args.date:
        t = All_data[3:4]
        txt_date = "".join(t)
        print(txt_date)
        
    elif args.performance:
        p = All_data[11:23]
        # print(p)
        temp_performace = []
        for i in p:
            # i = "".join(i.split())
            i = i.replace(" ", "")
            i = i.replace(":", "\t")
            # i = i.split(':')
            temp_performace.append(i)
        print(temp_performace)
        out_performance = "".join(temp_performace)
        print(out_performance)
        # excode = ast.literal_eval(out_performace)
    else:
        print("this is all experiment dcleaata you get from test, you can find it in your clipborad")
        print("txt中的实验数据为：")
        print("I(A)\tV(V)\tP(W)")
        # print(filecontents)
        str_filecontents = "".join(filecontents)
        print(str_filecontents)
        # 写入剪贴板
        writeclip(str_filecontents)
        
        
        
    # if args.time and args.performance:
    #    t = All_data[4:5]
    #    print(t)
    # elif args.time:
    #   print('time')
    # else args.perfperformance:
    #   print('performance')