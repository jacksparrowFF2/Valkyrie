#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   dark_IV.py
@Time    :   2019/12/28 11:00:37
@Author  :   SPH
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import os
import argparse
import win32clipboard as w
import win32con
import numpy as np
import xlwings as xw
import math

parser = argparse.ArgumentParser(description='该脚本用于对石墨烯IV数据进行处理')

parser.add_argument('-i', '--input_data', metavar='',
                    type=argparse.FileType(mode='r'))
parser.add_argument('-i2', '--input_area', metavar='',
                    type=argparse.FileType(mode='r'))
parser.add_argument('-e', '--excel', metavar = '', type = str, 
                    help = '保存数据的 excel 文件路径')
parser.add_argument('-n', '--factor', metavar='',
                    type=float)

parser.add_argument_group('基础选项')
parser.add_argument('-c', '--copy', action='store_true', help='将所有数据写入到剪贴板中')
parser.add_argument('-col', "--column", metavar = '', type=int, help='要提取的数据列')
parser.add_argument('-s1', '--step1', action='store_true', help='第一阶段数据处理')
parser.add_argument('-s2', '--step2', action='store_true', help='第二阶段数据处理')


groupA = parser.add_mutually_exclusive_group()
# parser.add_mutually_exclusive_group('高级选项')
groupA.add_argument('-L', '--Light', action='store_true', help='光照IV')
groupA.add_argument('-D', '--Dark', action='store_true', help='暗态IV')

groupB = parser.add_mutually_exclusive_group()
groupB.add_argument('-a', '--all', action='store_true', help='在剪贴板显示出当前txt文件的所有数据')
groupB.add_argument('-s', '--select', action='store_true', help='提取指定数据列到剪贴板')
groupB.add_argument('-wm', '--write_metadata', action='store_true', help='写入电压电流原始数据到excel文件中')
groupB.add_argument('-ws', '--write_statistics', action='store_true', help='写入电压电流统计数据到excel文件中')

groupC = parser.add_mutually_exclusive_group()
# parser.add_mutually_exclusive_group('高级选项')
groupC.add_argument('-C', '--creat', action='store_true', help='创建Excel文件')
groupC.add_argument('-f1', '--floor1', action='store_true', help='一楼IV测试设备')
groupC.add_argument('-f2', '--floor2', action='store_true', help='二楼IV测试设备')

args = parser.parse_args()


# Richard Constant(A cm^-2 k^-2)
RC = 252
# temperature(℃)
t = 25
# absolute temperature(k)
T = t + 273.15
# reverse bias saturation current density
J0 = 1
# Boltzmann Constant(J/k)
kB = 1.380649E-23
# electron 
elec = 1.6021766208E-19
# 构建复合因子
beta = elec/(kB*T)

def writeclip(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, astring)
    w.CloseClipboard()
def data_process(parameter_list):
    temp_list = []
    for i in parameter_list:
        x = '\t'.join(str(num) for num in i)
        # print(x)
        temp_list.append(x)
    print('输出第二阶段数据处理结果')
    # print(temp_list)
    out_list = []
    for i in temp_list:
        x = i + '\n'
        out_list.append(x)
    print('输出第三阶段数据处理结果')
    # print(out_list)
    out_str = ''.join(out_list)
    print('输出第三阶段数据处理结果')
    return out_str
def str2fstr(astring):
    in_str = astring.replace('\n', '')
    temp_list = in_str.split(':')
    i = temp_list[1].lstrip()
    return(i)

def str2list(astring):
    in_str = astring.replace('\n', '')
    temp_list = in_str.split('\t')
    out_list = []
    for i in temp_list:
        i = i.lstrip()
        out_list.append(i)
    return(out_list)

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
    if args.floor1:
        if args.Dark:
            if args.all:
                # 文件路径赋值给 infile
                infile = args.input_data
                # 从第 13 行处开始读取 txt 文件
                All_data = infile.readlines()
                filecontents = All_data[13:]
                print("this is all experiment dcleaata you get from test, you can find it in your clipborad")
                print("txt中的实验数据为：")
                print("I(A)\tV(V)\tP(mW)")
                # print(filecontents)
                str_filecontents = "".join(filecontents)
                print(str_filecontents)
                # 写入剪贴板
                writeclip(str_filecontents)
            elif args.select:
                # 文件路径赋值给 infile
                infile = args.input_data
                # 从第 20 行处开始读取 txt 文件
                All_data = infile.readlines()
                filecontents = All_data[13:]
                # print(filecontents)
                # print(type(filecontents))
                print("你选择输出的是第 %s 列" % (args.column))
                # 转化至程序排序方式
                n = args.column - 1
                print(n)
                # 去除可能存在的换行符
                while '\n' in filecontents:
                    filecontents.remove('\n')
                # 构建格式化列表
                out_select_list = selectcolumn_str(filecontents,n)
                # print(out_select_list)
                str_data = "\n".join(out_select_list)
                print(str_data)
                writeclip(str_data)
            elif args.write_metadata:
                # 文件路径赋值给 infile
                infile = args.input_data
                # 获取输入文件的项目编号
                item = os.path.split(str(infile))[1].split('.')[0].split('_')[0]
                print(item)
                print(type(item))
                # 面积文件路径赋值给 Ainfile
                Ainfile = args.input_area
                # 从第 2 行处开始读取 txt 文件
                A_All_data = Ainfile.readlines()
                Area = A_All_data[1:]
                # 构建面积列表
                A = selectcolumn_str(Area,2)
                A = list(map(float,A))
                print(A)
                # 构建定位标识
                A_index = selectcolumn_str(Area,0)
                A_index = list(map(int,A_index))
                print(A_index)
                # 找到输入文件对应的面积在列表中的位置
                p = A_index.index(int(item))
                print(p)
                # 输出对应的面积
                P = A[p]
                print(P)
                print(type(P))
                # 从第 13 行处开始读取 txt 文件
                All_data = infile.readlines()
                filecontents = All_data[13:]
                # print(filecontents)
                # 去除可能存在的换行符
                while '\n' in filecontents:
                    filecontents.remove('\n')
                while '\n' in Area:
                    filecontents.remove('\n')
                # 构建格式化列表-x
                x = selectcolumn_str(filecontents,0)
                # 构建格式化列表-y
                y = selectcolumn_str(filecontents,1)
                y = np.array(list(map(float,y)))
                print("开始填写excel")
                # 设定名称
                name = ["V(V)", "Jsc(mA/cm2)"+"-"+str(A_index[p])]
                name3 = ["V(V)", "Jsc(A/cm2)"+"-"+str(A_index[p])]
                # A/cm2
                y3 = list(map(str,list(abs(y/P*100))))
                # abs mA/cm2
                y2 = list(map(str,list(abs(y*1000/P*100))))
                # mA/cm2
                y = list(map(str,list(y*1000/P*100)))
                try:        
                    inexcel = args.excel
                    print('你输入的文件路径为：'+inexcel)
                    app = xw.App(visible=False,add_book=False)
                    wb = app.books.open(inexcel)
                    sht = wb.sheets['Dark I-V metadata']
                    sht2 = wb.sheets['ABSDark I-V metadata']
                    sht3 = wb.sheets['A-ABSDark I-V metadata']
                    # 获取表格坐标信息
                    info = sht.range('A1').expand('table')
                    row = info.last_cell.row
                    col = info.last_cell.column
                    # 计算出要添加的一列位置
                    coll = col + 1
                    str_coll = str(coll)
                    print('数据添加所在列：'+str_coll)
                    str_col = str(col)
                    print('原表格最后一列：'+str_col)
                    # 输出结果
                    if col == 1:
                        # 填写坐标名称
                        sht.range((1,col),(1,coll)).options(transpose = False).value = name
                        sht2.range((1,col),(1,coll)).options(transpose = False).value = name
                        sht3.range((1,col),(1,coll)).options(transpose = False).value = name3
                        # 填写X坐标数据
                        sht.range((2,col),(2+len(x),col)).options(transpose = True).value = x
                        sht2.range((2,col),(2+len(x),col)).options(transpose = True).value = x
                        sht3.range((2,col),(2+len(x),col)).options(transpose = True).value = x
                        # 填写y坐标数据
                        sht.range((2,coll),(2+len(x),coll)).options(transpose = True).value = y
                        sht2.range((2,coll),(2+len(x),coll)).options(transpose = True).value = y2
                        sht3.range((2,coll),(2+len(x),coll)).options(transpose = True).value = y3
                        print('注入完成')
                    else:
                        # 填写坐标名称
                        sht.range((1,coll),(1,coll)).options(transpose = False).value = name[1]
                        sht2.range((1,coll),(1,coll)).options(transpose = False).value = name[1]
                        sht3.range((1,coll),(1,coll)).options(transpose = False).value = name[1]
                        # 填写y坐标数据
                        sht.range((2,coll),(2+len(x),coll)).options(transpose = True).value = y
                        sht2.range((2,coll),(2+len(x),coll)).options(transpose = True).value = y2
                        sht3.range((2,coll),(2+len(x),coll)).options(transpose = True).value = y3
                        print('注入完成')
                    print('实验数据注入完成！')
                finally:
                    if wb:
                        wb.save()
                        wb.close()
                        app.kill()
            else:
                print("请选择处理模式")
            """ if args.copy:
                infile = args.input
                All_data = infile.readlines()
                # print(All_data)
                Area = float(All_data[7].split('\t')[1].split('\n')[0])
                V = []
                I = []
                J = []
                for line in All_data[13:]:
                    if line != '\n':
                        V1,I1 = line.split('\t')
                        # print(V1,I1)
                        V.append(float(V1))
                        I.append(float(I1))
                # print(V)
                # print(I)
                # 构建 J 数列
                J = list(map(lambda x : x/Area, I))
                # 构建 1/(J*1000)
                J_1000 = list(map(lambda x: 1/(x*1000), J))
                # 构建对数处理 J 数列
                absJ = list(map(abs, J))
                # 进行对数处理
                lnJ = list(map(math.log, absJ))
                
                if args.step1:
                    name = ['V', 'I', 'J', '1/J_1000', 'absJ', 'lnJ']
                    Unit = ['V', 'A', 'A/cm2', 'A-1cm-2', 'A/cm2', 'A/cm2']
                    temp_list1 = [list(item) for item in zip(V, I, J, J_1000, absJ,  lnJ)]
                    temp_list1.insert(0,'')
                    temp_list1.insert(0,Unit)
                    temp_list1.insert(0,name)
                    # print(temp_list1)
                    out_str = data_process(temp_list1)
                    # print(out_str)
                    writeclip(out_str)
                elif args.step2:
                    # 创建HJ数列
                    HJ_temp = list(map(lambda x: args.factor*math.log(x/(RC*T**2))/beta, J))
                    HJ = list(np.array(V) - np.array(HJ_temp))
                    name = ['HJ']
                    Unit = ['A/cm2']
                    temp_list1 = [list(item) for item in zip(HJ)]
                    # print(temp_list1)
                    temp_list1.insert(0,'')
                    temp_list1.insert(0,Unit)
                    temp_list1.insert(0,name)
                    out_str = data_process(temp_list1)
                    # print(out_str)
                    writeclip(out_str)
                else:
                    print('请选择处理阶段') """
        if args.Light:
            if args.all:
                # 文件路径赋值给 infile
                infile = args.input_data
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
            elif args.select:
                # 文件路径赋值给 infile
                infile = args.input_data
                # 从第 20 行处开始读取 txt 文件
                All_data = infile.readlines()
                filecontents = All_data[22:]
                # print(filecontents)
                # print(type(filecontents))
                print("你选择输出的是第 %s 列" % (args.column))
                # 转化至程序排序方式
                n = args.column - 1
                print(n)
                # 去除可能存在的换行符
                while '\n' in filecontents:
                    filecontents.remove('\n')
                # 构建格式化列表
                out_select_list = selectcolumn_str(filecontents,n)
                # print(out_select_list)
                str_data = "\n".join(out_select_list)
                print(str_data)
                writeclip(str_data)
            elif args.write_metadata:
                # 文件路径赋值给 infile
                infile = args.input_data
                # 获取输入文件的项目编号
                item = os.path.split(str(infile))[1].split('.')[0].split('_')[0]
                print(item)
                print(type(item))
                # 面积文件路径赋值给 Ainfile
                Ainfile = args.input_area
                # 从第 2 行处开始读取 txt 文件
                A_All_data = Ainfile.readlines()
                Area = A_All_data[1:]
                # 构建面积列表
                A = selectcolumn_str(Area,2)
                A = list(map(float,A))
                print(A)
                # 构建定位标识
                A_index = selectcolumn_str(Area,0)
                A_index = list(map(int,A_index))
                print(A_index)
                # 找到输入文件对应的面积在列表中的位置
                p = A_index.index(int(item))
                print(p)
                # 输出对应的面积
                P = A[p]
                print(P)
                print(type(P))
                # 从第 22 行处开始读取 txt 文件
                All_data = infile.readlines()
                filecontents = All_data[22:]
                # print(filecontents)
                # 去除可能存在的换行符
                while '\n' in filecontents:
                    filecontents.remove('\n')
                while '\n' in Area:
                    filecontents.remove('\n')
                # 构建格式化列表-x
                x = selectcolumn_str(filecontents,0)
                # 构建格式化列表-y
                y = selectcolumn_str(filecontents,1)
                y = np.array(list(map(float,y)))
                print("开始填写excel")
                try:        
                    inexcel = args.excel
                    print('你输入的文件路径为：'+inexcel)
                    app = xw.App(visible=False,add_book=False)
                    wb = app.books.open(inexcel)
                    sht = wb.sheets['Light I-V metadata']
                    # 获取表格坐标信息
                    info = sht.range('A1').expand('table')
                    row = info.last_cell.row
                    col = info.last_cell.column
                    # 计算出要添加的一列位置
                    coll = col + 1
                    str_coll = str(coll)
                    print('数据添加所在列：'+str_coll)
                    str_col = str(col)
                    print('原表格最后一列：'+str_col)
                    y = list(map(str,list(y/P*100)))
                    # 设定名称
                    name = ["V(V)", "Jsc(mA/cm2)"+"-"+str(A_index[p])]
                    # 输出结果
                    if col == 1:
                        # 填写坐标名称
                        sht.range((1,col),(1,coll)).options(transpose = False).value = name
                        # 填写X坐标数据
                        sht.range((2,col),(2+len(x),col)).options(transpose = True).value = x
                        # 填写y坐标数据
                        sht.range((2,coll),(2+len(x),coll)).options(transpose = True).value = y
                        print('注入完成')
                    else:
                        # 填写坐标名称
                        sht.range((1,coll),(1,coll)).options(transpose = False).value = name[1]
                        # 填写y坐标数据
                        sht.range((2,coll),(2+len(x),coll)).options(transpose = True).value = y
                        print('注入完成')
                    print('实验数据注入完成！')
                finally:
                    if wb:
                        wb.save()
                        wb.close()
                        app.kill()
            elif args.write_statistics:
                # 文件路径赋值给 infile
                infile = args.input_data
                # 从第 9 行处开始读取 txt 文件
                All_data = infile.readlines()
                filecontents = All_data[9:]
                # print(filecontents)
                # 获取编号
                Material = str2list(All_data[4:5][0])[1]
                print(Material)                
                # 面积文件路径赋值给 Ainfile
                Ainfile = args.input_area
                # 从第 2 行处开始读取 txt 文件
                A_All_data = Ainfile.readlines()
                Area = A_All_data[1:]
                # 构建面积列表
                A = selectcolumn_str(Area,2)
                A = list(map(float,A))
                print(A)
                # 构建定位标识
                A_index = selectcolumn_str(Area,0)
                A_index = list(map(int,A_index))
                print(A_index)
                # 找到输入文件对应的面积在列表中的位置
                p = A_index.index(int(Material))
                print(p)
                # 输出对应的面积
                P = A[p]
                print(P)
                print(type(P))
                # 去除可能存在的换行符
                while '\n' in filecontents:
                    filecontents.remove('\n')
                while '\n' in Area:
                    filecontents.remove('\n')
                # 构建面积列表
                A = selectcolumn_str(Area,2)
                print(A)
                # 构建格式化列表-y
                y = str2list(filecontents[1])
                y[0] = Material
                print(y)
                try:        
                    inexcel = args.excel
                    print('你输入的文件路径为：'+inexcel)
                    app = xw.App(visible=False,add_book=False)
                    wb = app.books.open(inexcel)
                    sht = wb.sheets['statistics metadata']
                    # 获取表格坐标信息
                    info = sht.range('A1').expand('table')
                    row = info.last_cell.row
                    col = info.last_cell.column
                    print('原表格最后一行：'+str(row))
                    # 计算出要添加的一行位置
                    rowl = row + 1
                    print('数据添加所在行：'+str(rowl))
                    # 计算倍率
                    n = 0.45/(P/100)
                    # 计算eff矫正效率
                    eff = str(float(y[-1])*n)
                    print(eff)
                    # 计算Jsc矫正电流
                    Jsc = str(float(y[-3])*n)
                    # 构建最终数据列
                    y = y + [str(n),str(Jsc),str(eff)]
                    print(y)
                    print("开始填写excel")
                    # 输出结果
                    # 填写y坐标数据
                    sht.range('A'+str(rowl),'P'+str(rowl)).value = y
                    print('注入完成')
                    print('实验数据注入完成！')
                finally:
                    if wb:
                        wb.save()
                        wb.close()
                        app.kill()
            else:
                print("请选择处理模式")
    elif args.floor2:
        if args.dark:
            if args.copy:
                infile = args.input
                All_data = infile.readlines()
                # print(All_data)
                index = []
                V = []
                I = []
                Power = []
                J = []
                for line in All_data[1:]:
                    index1, V1, I1, Power1, J1 = line.split()
                    index.append(float(index1))
                    V.append(float(V1))
                    I.append(float(I1))
                    Power.append(float(Power1))
                    J.append(float(J1))
                    # index= line.split()
                # 构建对数处理 I 数列
                I = list(map(abs, I))
                # 进行对数处理
                lnI = list(map(math.log, I))
                Power = list(map(abs, Power))
                # 构建对数处理 J 数列
                J = list(map(abs, J))
                # 进行对数处理
                lnJ = list(map(math.log, J))
                J_1000 = list(map(lambda x: x*1000, J))
                # 构建正常 J 数组，单位为 mA/cm2
                rJ_1000 = list(map(lambda x: -x*1000, J))
                # print(HJ)
                if args.step1:
                    name = ['Voltage', 'I', 'Power', 'J', 'lnJ', 'J_1000', 'rJ_1000']
                    Unit = ['V', 'A', 'w', 'A/cm2', 'A/cm2', 'mA/cm2', 'mA/cm2']
                    temp_list1 = [list(item) for item in zip(V, I, Power, J, lnJ, J_1000, rJ_1000)]
                    temp_list1.insert(0,'')
                    temp_list1.insert(0,Unit)
                    temp_list1.insert(0,name)
                    # print(temp_list1)
                    out_str = data_process(temp_list1)
                    print(out_str)
                    writeclip(out_str)

                    # temp_list2 = []
                    # for i in temp_list1:
                    #     x = '\t'.join(str(num) for num in i)
                    #     # print(x)
                    #     temp_list2.append(x)
                    # print('输出第二阶段数据处理结果')
                    # # print(temp_list2)
                    # out_list = []
                    # for i in temp_list2:
                    #     x = i + '\n'
                    #     out_list.append(x)
                    # print('输出第三阶段数据处理结果')
                    # # print(out_list)
                    # out_str = ''.join(out_list)
                    # print('输出第三阶段数据处理结果')
                    # # print(out_str)
                    # writeclip(out_str)
                elif args.step2:
                    # 创建HJ数列
                    HJ_temp = list(map(lambda x: args.factor*math.log(x/(RC*T**2))/beta, J))
                    HJ = list(np.array(V) - np.array(HJ_temp))
                    name = ['HJ']
                    Unit = ['A/cm2']
                    temp_list1 = [list(item) for item in zip(HJ)]
                    # print(temp_list1)
                    temp_list1.insert(0,'')
                    temp_list1.insert(0,Unit)
                    temp_list1.insert(0,name)
                    out_str = data_process(temp_list1)
                    # print(out_str)
                    writeclip(out_str)
                else:
                    print('请选择处理阶段')
        print('处理结束')
    elif args.creat:
        print("开始创建excel")
        name = ['Material', 'Time/s', 'Serial NO.', 'Voc/V', 'Isc/mA', 'Pmax/mW', 'Vpmax/V', 'Ipmax/mA', 'Rs/ohm', 'Rsh/ohm', 'Jsc/mA.cm-2', 'FF', 'η/%', 'n', 'Jsc','eff']
        try:
            app = xw.App(visible=False,add_book=False)
            # wb = app.books.add()
            wb = app.books.add()
            wb.sheets["sheet1"].name = "A-ABSDark I-V metadata"
            wb.sheets.add("sheet2")
            wb.sheets["sheet2"].name = "ABSDark I-V metadata"
            wb.sheets.add("sheet3")
            wb.sheets["sheet3"].name = "Dark I-V metadata"
            wb.sheets.add("sheet4")
            wb.sheets["sheet4"].name = "Light I-V metadata"
            wb.sheets.add("sheet5")
            wb.sheets["sheet5"].name = "statistics metadata"
            sht = wb.sheets['statistics metadata']
            sht.range('A1','P1').value = name
            # 格式化
            # 对表格进行美化
            # 对第一行标题进行格式化
            sht.range('A1').expand('right').api.HorizontalAlignment = -4108
            sht.range('A1').expand('right').api.VerticalAlignment = -4108
            # 行高
            sht.api.Rows(1).RowHeight = 20
            # 列宽
            sht.api.Columns("A:P").Columnwidth = 15
            print('格式化完成')
        finally:
            if wb:
                wb.save(args.excel)
                wb.close()
                app.kill()
    else:
        print('请选择测试设备所在楼层、处理模式、数据处理阶段')
