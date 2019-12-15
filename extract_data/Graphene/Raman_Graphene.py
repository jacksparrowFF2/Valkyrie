#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   Raman.py
@Time    :   2019/09/29 13:45:41
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import os
from operator import itemgetter
import win32con
import win32clipboard as w
import xlwings as xw
import argparse

# 创建命令解释器
parser = argparse.ArgumentParser('该脚本旨在帮助你更方便的对拉曼数据进行整理')

parser.add_argument('-i', '--input', metavar = '', type = argparse.FileType(mode='r'), 
                    help = '要进行整理的拉曼数据文件路径')
parser.add_argument('-e', '--excel', metavar = '', type = str, 
                    help = '保存数据的 excel 文件路径')

group = parser.add_argument_group('基础选项')
group.add_argument('-c', '--column', metavar = '', type = int, help = '要提取的数据列')

group = parser.add_argument_group('进阶选项')
group.add_argument('-wc', '--wconditon', action = 'store_true', 
                   help = '模式：将实验条件写入指定 excel 表格')
group.add_argument('-wr', '--wraman', action = 'store_true', 
                   help = '模式：将拉曼数据写入指定 excel 表格')
group.add_argument('-wf','--wfit', action = 'store_true', 
                   help = '模式：将拟合结果写入指定 excel 表格') 
group.add_argument('-cf','--cfit', action = 'store_true', 
                   help = '模式：将拟合结果写入到剪贴板') 
group.add_argument('-cr','--copyselect', action = 'store_true', 
                   help = '模式：将提取的数据列写入剪贴板') 
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

def fitformat(astring):
    a = astring.split('\n')
    while '' in a:
        a.remove('')
    c = []
    for item in a:
        b = item.split('\t')
        c.append(b)
    d = sorted(c,key=itemgetter(5))
    return(d)

def conformat(astring):
    a = astring.split('\n')
    b = []
    for i in a:
        i = i.replace('\r', '')
        b.append(i.split(':'))
    c = {}
    for i in b:
        c[i[0]] = i[1]
    return(c)

def str2list(astring):
    n = args.column - 1
    a = astring.split('\n')
    while '' in a:
      a.remove('')
    b = []
    for i in a:
        b.append(i.split('\t')[n])
    return(b)
if __name__ == '__main__':
    if args.wconditon:
        print('wcondition')
        excode = conformat(getclip())
        print(excode)
        # 实验数据变量赋值
        date = excode["日期"]
        power = int(excode["初始输入功率(w)"])-int(excode["初始反馈功率(w)"])
        Ar = int(excode["Ar(sccm)"])
        H2 = int(excode["H2(sccm)"])
        CH4 = int(excode["CH4(sccm)"])
        pressure = int(excode["压强(pa)"])
        temp = int(excode["温度(℃)"])
        sub1 = excode["衬底1"]
        sub2 = excode["衬底2"]
        metaltype = excode["金属网"]
        note = excode["实验目的"]
        time = int(excode["持续时间(min)"])
        SR = str(excode["方阻(kΩ/□)"])
        try:
            # 开始对 excel 文件进行操作
            inexcel = args.excel
            # 创建 App 进程
            app = xw.App(visible = False, add_book = False)
            # 链接工作表
            wb = app.books.open(inexcel)
            # 对指定工作表进行编辑
            sht = wb.sheets['Ratio Metadata']
            # 获取当前EXCEL表格的行数与列数
            info = sht.range('A1').expand('table')
            row = info.last_cell.row
            col = info.last_cell.column
            # 计算出要添加的一行位置
            rowl = row + 1
            # # 调试输出
            # print('原表格最后一行：'+str(row))
            # print('数据添加所在行：'+str(rowl))
            # 注入EXCEL公式
            # 创建实验条件数据列
            excondition = [note+"+"+sub1+"+"+sub2, metaltype, Ar, H2, CH4, time, power, pressure, temp, SR]
            # 注入EXCEL公式
            sht.range('A'+str(rowl)).formula = '=B%s&CHAR(10)&AG%s'%(rowl,rowl)
            sht.range('K'+str(rowl)).formula = '=C%s/E%s'%(rowl,rowl)
            sht.range('L'+str(rowl)).formula = '=I%s/E%s'%(rowl,rowl)
            sht.range('M'+str(rowl)).formula = '=C%s/G%s'%(rowl,rowl)
            sht.range('N'+str(rowl)).formula = '=IF(88>J%s,45/(88-J%s),"bulk")'%(rowl,rowl)
            sht.range('Y'+str(rowl)).formula = '=Q%s*1.415'%(rowl)
            sht.range('Z'+str(rowl)).formula = '=R%s*1.01'%(rowl)
            sht.range('AA'+str(rowl)).formula = '=S%s*0.719'%(rowl)
            sht.range('AB'+str(rowl)).formula = '=Q%s&"/"&R%s&"/"&S%s'%(rowl,rowl,rowl)
            sht.range('AC'+str(rowl)).formula = '=Y%s/SUM(Y%s+Z%s+AA%s)'%(rowl,rowl,rowl,rowl)
            sht.range('AD'+str(rowl)).formula = '=Z%s/SUM(Y%s+Z%s+AA%s)'%(rowl,rowl,rowl,rowl)
            sht.range('AE'+str(rowl)).formula = '=AA%s/SUM(Y%s+Z%s+AA%s)'%(rowl,rowl,rowl,rowl)
            sht.range('AF'+str(rowl)).formula = '=AA%s/Z%s'%(rowl,rowl)
            sht.range('AG'+str(rowl)).formula = '=P%s&"/"&Q%s&"/"&R%s&"/"&S%s&"/"&T%s&"/"&U%s&"/"&V%s&"/"&W%s'%(rowl,rowl,rowl,rowl,rowl,rowl,rowl,rowl)
                # 注入实验条件数据
            sht.range('O'+str(rowl),'X'+str(rowl)).value = excondition
            sht.range('B'+str(rowl)).value = date
            # 对表格进行美化
            #     # 对第一行标题进行格式化
            # sht.range('A1').expand('right').api.HorizontalAlignment = -4108
            # sht.range('A1').expand('right').api.VerticalAlignment = -4108
                # 对数据列 A 进行格式化为水平居中对齐，字体加粗
            sht.range('A2').expand('down').api.HorizontalAlignment = -4152
            sht.range('A2').expand('down').api.VerticalAlignment = -4108
            sht.range('A2').expand('down').api.font.Bold = True
                # 对数据列 A 进行自动换行
            sht.range('A2').expand('down').api.WrapText = True
                # 对数据列 B 进行格式化为垂直居中对齐，水平靠右，加粗
            sht.range('B2').expand('down').api.HorizontalAlignment = -4152
            sht.range('B2').expand('down').api.VerticalAlignment = -4108
            sht.range('B2').expand('down').api.font.Bold = True
                # 对数据列 C:AE 进行格式化为水平垂直居中对齐
            sht.range('C2:AE'+str(rowl)).api.HorizontalAlignment = -4108
            sht.range('C2:AE'+str(rowl)).api.VerticalAlignment = -4108
                # 对数据列 B 进行格式化为垂直居中对齐，水平靠右，加粗
            sht.range('AF2').expand('down').api.HorizontalAlignment = -4152
            sht.range('AF2').expand('down').api.VerticalAlignment = -4108
                # 将数据列 AC:AE 数据显示格式化为百分比
            sht.range('AC2:AE2').expand('down').api.style = "Percent"
                # 将数据列 C:M,Y:AA 数据显示格式化为保留两位小数点
            sht.range('C2:J'+str(rowl)).api.NumberFormat = "0.00_);(0.00)"
            sht.range('K2:N'+str(rowl)).api.NumberFormat = "0.00_);(0.00)"
            sht.range('Y2:AA2').expand('down').api.NumberFormat = "0.00_);(0.00)"
                # 将数据列 AF 数据显示格式化为保留一位小数点
            sht.range('AF2').expand('down').api.NumberFormat = "##.0_)"
                # 将数据列 Q 数据显示格式化为 3 位
            sht.range('Q2').expand('down').api.NumberFormat = "000"
                # 将数据列 R:S 数据显示格式化为 2 位
            sht.range('R2:S2').expand('down').api.NumberFormat = "00"
                # 将数据列 T 数据显示格式化为 3 位
            sht.range('T2').expand('down').api.NumberFormat = "000"
            # 格式化完成提示
            print('EXCEL 格式化完成')
        finally:
            if wb:
                # 保存文件
                wb.save()
                # 关闭文件
                wb.close()
                # 结束进程
                app.kill
    elif args.wraman:
        print('wraman')
        infile = args.input
        in_ramandata = infile.read()
        print("你选择输出的数据列为:%s" %(args.column))
        out_list = str2list(in_ramandata)
        # print(out_list)
        # 构建间接引用公式
        part1 = '=INDIRECT('
        part2 = '"'
        part3 = '\'Ratio Metadata\''
        part4 = '!$A'
        part5  = '"&COLUMN())'
        formula = part1 + part2 + part3 + part4 + part5
        print(formula)
        # 将公式插入到列表中
        out_list.insert(0,formula)
        # print(out_list)
        try:
            # 开始对 excel 操作
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
            # # 调试输出
            # print(info)
            row = info.last_cell.row
            col = info.last_cell.column
            # 计算出要添加的一行位置
            coll = col + 1
            str_coll = str(coll)
            print('数据添加所在列：'+str_coll)
            str_col = str(col)
            print('原表格最后一列：'+str_col)
            # 注入实验数据
            sht.range((1,coll),(row,coll)).options(transpose = True).value = out_list
            # 输出结果
            print('实验数据注入完成！')

            # 对新增列进行格式化
            # 对新增首行单元格进行格式化
            sht.range((1,coll)).api.HorizontalAlignment= -4152
            sht.range((1,coll)).api.VerticalAlignment= -4108
            sht.range((1,coll)).api.WrapText= True
            # 对新增数据单元格进行格式化
            sht.range((2,coll),(row,coll)).api.HorizontalAlignment= -4108
            sht.range((2,coll),(row,coll)).api.VerticalAlignment= -4108
            # 将新增单元格的列宽调整为 40
            sht.api.columns(coll).ColumnWidth = 40
        finally:
            if wb:
                #保存文件
                wb.save()
                # 关闭文件
                wb.close()
                # 结束进程
                app.kill
    elif args.copyselect:
        print('copyselect')
        infile = args.input
        in_ramandata = infile.read()
        print("你选择输出的数据列为:%s" %(args.column))
        out_str = "\n".join(str2list(in_ramandata))
        print(out_str)
        writeclip(out_str)
    elif args.wfit:
        print('wfit')
        temp_list = fitformat(getclip())
        out_list = []
        for i in temp_list:
            out_list.append(i[2])
            out_list.append(i[3])
        print(out_list)
        del out_list[8:10]
        # 创建拟合结果数据列
        fitdata = out_list
        try:
            # 开始对 excel 文件进行操作
            inexcel = args.excel
            # 创建 App 进程
            app = xw.App(visible = False, add_book = False)
            # 链接工作表
            wb = app.books.open(inexcel)
            # 对指定工作表进行编辑
            sht = wb.sheets['Ratio Metadata']
            # 获取当前EXCEL表格的行数与列数
            info = sht.range('A1').expand('table')
            row = info.last_cell.row
            col = info.last_cell.column
            # # 调试输出
            # print('原表格最后一行：'+str(row))
            # print('数据添加所在行：'+str(rowl))
            # 注入实验条件数据
            sht.range('C'+str(row), 'J'+str(row)).value = fitdata
        finally:
            if wb:
                # 保存文件
                wb.save()
                # 关闭文件
                wb.close()
                # 结束进程
                app.kill
    elif args.cfit:
        print('cfit')
        temp_list = fitformat(getclip())
        out_list = []
        for i in temp_list:
            out_list.append(i[2]+'\t')
            out_list.append(i[3]+'\t')
        out_list[7] = out_list[7].replace('\t', '')
        print(out_list[7])
        del out_list[8:10]
        print(out_list)
        out_str = "".join(out_list)
        print(out_str)
        writeclip(out_str)
    else:
        print('请输入 -h 以查看使用说明')



input("SSS")
