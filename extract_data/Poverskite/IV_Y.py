#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   IV.py
@Time    :   2019/10/09 20:57:31
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib

# 导入系统组件
import os
# 导入 EXCEL 操作
import xlwings as xw
# 导入剪贴板组件
import win32clipboard as w
import win32con
# 导入命令行参数组件
import argparse
# 创建命令解释器
parser = argparse.ArgumentParser(
    description='从 txt 文件中提取 I-V 数据')
# 创建命令行输入参数，输入参数为文件路径
parser.add_argument('-i', '--input', metavar = '', type=argparse.FileType(mode='r'),
                    help='需要提取数据的 txt 文件路径')
# 创建附属命令行参数，增加可选输出第二列的选项
group = parser.add_argument_group(description = '基础选项')
group.add_argument('-c', "--column", metavar = '', type=int,
                    help='要提取的数据列')
group.add_argument('-e', "--excel", metavar = '', type=str,
                    help='写入数据的 excel 文件路径')
# 创建互斥锁
group = parser.add_argument_group('高级选项')
# group = parser.add_mutually_exclusive_group(description = 'Basic options')
group.add_argument('-a', '--all', action='store_true',
                   help='在剪贴板显示出当前txt文件的所有数据')
group.add_argument('-s', '--select', action='store_true',
                   help='提取指定数据列到剪贴板')
group.add_argument('-d', '--delete', action='store_true',
                   help='删除指定数据列并将剩余数据列提取到剪贴板')
group.add_argument('-w', '--write', action='store_true',
                   help='将数据文件写入至指定的Excel文件')
args = parser.parse_args()

# 创建剪贴板写入函数
def writeclip(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, astring)
    w.CloseClipboard()

def str2fstr(astring):
    in_str = astring.replace('\n', '')
    temp_list = in_str.split(':')
    i = temp_list[1].lstrip()
    return(i)

def ftime(astring):
    in_str = astring.replace('\n', '')
    temp_list = in_str.split(' ')
    i = temp_list[3] + ' ' + temp_list[4]
    return(i)

def selectcolumn_str(astring,column):
    out_select_list  = []
    for i in astring:
        templist = i.replace('\n','').split('\t')
        out_select_list.append(templist[column])
    return(out_select_list)

def selectcolumn_del(astring,column):
    out_select_list  = []
    for i in astring:
        templist = i.replace('\n','').split('\t')
        templist.pop(column)
        tempstr = "\t".join(templist)
        out_select_list.append(tempstr)
        out_str = "\n".join(out_select_list) 
    return(out_str)

if __name__ == '__main__':
    if args.select:
        # 文件路径赋值给 infile
        infile = args.input
        # 从第 30 行处开始读取 txt 文件
        All_data = infile.readlines()
        filecontents = All_data[22:]
        print("你选择输出的是第 %s 列" % (args.column))
        # 转化至程序排序方式
        n = args.column - 1
        # 构建格式化列表
        out_select_list = selectcolumn_str(filecontents,n)
        print(out_select_list)
        str_data = "\n".join(out_select_list)
        print(str_data)
        writeclip(str_data)
    elif args.delete:
        # 文件路径赋值给 infile
        infile = args.input
        # 从第 30 行处开始读取 txt 文件
        All_data = infile.readlines()
        filecontents = All_data[30:]
        print("the remaining data as follow show:")
        # 转化至程序排序方式
        n = args.column - 1
        remaining_list = selectcolumn_del(filecontents,n)
        print(remaining_list)
        writeclip(remaining_list)
    elif args.all:
        # 文件路径赋值给 infile
        infile = args.input
        # 从第 22 行处开始读取 txt 文件
        All_data = infile.readlines()
        filecontents = All_data[22:]
        print("this is all experiment dcleaata you get from test, you can find it in your clipborad")
        print("txt中的实验数据为：")
        print("I(A)\tV(V)\tP(mW)")
        # print(filecontents)
        str_filecontents = "".join(filecontents)
        print(str_filecontents)
        # 写入剪贴板
        writeclip(str_filecontents)
    elif args.write:
        # 文件路径赋值给 infile
        infile = args.input
        # 获取文件名
        filename = os.path.split(str(infile))[1].split('.')[0]
        All_data = infile.readlines()
        # 时间
        time = ftime(All_data[3])
        print(time)
        # 面积
        area = str2fstr(All_data[9])
        print(area)
        # 从第 11:22 行处开始读取 txt 文件
        in_performance = All_data[11:23]
        # 调试输出
        print(in_performance)
        out_performance = []
        for i in in_performance:
            i =  str2fstr(i)
            out_performance.append(i)
        # 性能数据
        print(out_performance)
        refine_data = []
        refine_data = [filename,
                       out_performance[5],
                       out_performance[0],
                       out_performance[7],
                       out_performance[6],
                       out_performance[-4],
                       out_performance[-3]
                       ]
        print(refine_data) 
        filecontents = All_data[30:]
        # 构建格式化列表
        I = selectcolumn_str(filecontents,0)
        V = selectcolumn_str(filecontents,1)
        P = selectcolumn_str(filecontents,2)
        I.insert(0,filename+'_I')
        V.insert(0,filename+'_V')
        P.insert(0,filename+'_P')
        print(I)
        print(V)
        print(P)
        try:
            inexcel = args.excel
            print('你输入的文件路径为：'+inexcel)
            # 开始对 EXCEL 文件进行编辑
            app = xw.App(visible = False,add_book = False)
            # 打开指定的 EXCEL 文件
            wb = app.books.open(inexcel)
            # 链接工作表指定工作表
            sht = wb.sheets['I-V Performance']
            sht2 = wb.sheets['refine data']
            sht3 = wb.sheets['raw data']
            # 获取表格尺寸
            info = sht.range('A1').expand('table')
            print(info)
            # 计算出数据要添加的位置
            row = info.last_cell.row
            col = info.last_cell.column
            rowl = row + 1
            print('原表格最后一行:'+str(row))
            print('数据添加所在行:'+str(rowl))
            # 向表格 1 注入性能数据
            sht.range('A'+str(rowl)).value = filename
            sht.range('B'+str(rowl)).value = area
            sht.range('C'+str(rowl),'N'+str(rowl)).value = out_performance
            # 对表格 1 进行格式化
                # 水平居中对齐
            sht.range('A'+str(rowl)).expand('right').api.HorizontalAlignment = -4108
            sht.range('A'+str(rowl)).expand('right').api.VerticalAlignment = -4108
                # 调整单元格的宽
            sht.api.columns(1).ColumnWidth = 46.56
                # 调整单元格行高为 30
            sht.api.Rows(rowl).RowHeight = 30
            # 格式化完成提示
            print('表格 1 格式化完成')
            
            # 向表格 2 注入性能数据
            # 获取表格尺寸
            info2 = sht2.range('A1').expand('table')
            print(info2)
            # 计算出数据要添加的位置
            row2 = info2.last_cell.row
            col2 = info2.last_cell.column
            rowl2 = row2 + 1
            print('表格2最后一行:'+str(row2))
            print('数据添加所在行:'+str(rowl2))
            sht2.range('A'+str(rowl2),'G'+str(rowl2)).value = refine_data
            # 对表格 2 进行格式化
                # 水平居中对齐
            sht2.range('A'+str(rowl2)).expand('right').api.HorizontalAlignment = -4108
            sht2.range('A'+str(rowl2)).expand('right').api.VerticalAlignment = -4108
                # 调整数字格式
            sht2.range('D'+str(rowl2)).expand('right').api.NumberFormat = "##.0_)"
                # 调整单元格行高为 30
            sht.api.Rows(rowl2).RowHeight = 30
            # 格式化完成提示
            print('表格 2 格式化完成')
            
            # 向表格 3 注入原始数据
            # 获取表格尺寸
            info3 = sht3.range('A1').expand('table')
            print(info3)
            # 计算出数据要添加的位置
            row3 = info3.last_cell.row
            col3 = info3.last_cell.column
            coll3 = col3 + 1
            print('表格3最后一列:'+str(col3))
            print('数据添加所在列:'+str(coll3))
            if col3 == 1:
                sht3.range((1,col3),(row3,col3)).options(transpose = True).value = I
                sht3.range((1,coll3),(row3,coll3)).options(transpose = True).value = V
                sht3.range((1,coll3+1),(row3,coll3+1)).options(transpose = True).value = P
                # 对表格 3 进行格式化
                    # 对新增数据单元格进行格式化
                sht3.range((1,col),(row3,coll3+1)).api.HorizontalAlignment= -4108
                sht3.range((1,col),(row3,coll3+1)).api.VerticalAlignment= -4108
                    # 水平居中对齐
                sht3.range('A1').expand('table').api.HorizontalAlignment = -4108
                sht3.range('A1').expand('table').api.VerticalAlignment = -4108
                    # 设定单元格宽度
                sht3.api.columns(col3).ColumnWidth = 20
                sht3.api.columns(coll3).ColumnWidth = 20
                sht3.api.columns(coll3+1).ColumnWidth = 20
            else:
                sht3.range((1,coll3),(row3,coll3)).options(transpose = True).value = I
                sht3.range((1,coll3+1),(row3,coll3+1)).options(transpose = True).value = V
                sht3.range((1,coll3+2),(row3,coll3+2)).options(transpose = True).value = P
                # 对表格 3 进行格式化
                    # 对新增数据单元格进行格式化
                sht3.range((1,coll3),(row3,coll3+2)).api.HorizontalAlignment= -4108
                sht3.range((1,coll3),(row3,coll3+2)).api.VerticalAlignment= -4108
                    # 设定单元格宽度
                sht3.api.columns(coll3).ColumnWidth = 20
                sht3.api.columns(coll3+1).ColumnWidth = 20
                sht3.api.columns(coll3+2).ColumnWidth = 20
            # 格式化完成提示
            print('表格 3 格式化完成')
            
        finally:
            if wb:
                # 保存文件
                wb.save()
                # 关闭文件
                wb.close()
                # 关闭进程
                app.kill()   
    else:
        print("请输入 -h 以查看帮助")