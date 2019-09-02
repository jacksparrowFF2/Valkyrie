#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   add_data_ver2.py
@Time    :   2019/09/02 11:11:22
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, Liugroup-NLPR-CASIA
@Desc    :   None
'''

# here put the import lib


import win32clipboard as wc
import win32con
import ast
import xlwings as xw
import get_table_info
from excel_formula import  A,J,K,L,X,Y,Z,AA,AB,AC
from get_table_info import app,wb,sht,info,row,rowl

#获取剪贴板内容
def getCopyText():
    wc.OpenClipboard()
    copy_text = wc.GetClipboardData(win32con.CF_UNICODETEXT)
    wc.CloseClipboard
    return copy_text

#将字符串类型转变为字典类型

excode = ast.literal_eval(getCopyText())
#输出变量类型
print(type(excode))

date = excode["日期"]
power = int(excode["初始输入功率(w)"])-int(excode["初始反馈功率(w)"])
Ar = int(excode["Ar(sccm)"])
H2 = int(excode["H2(sccm)"])
CH4 = int(excode["CH4(sccm)"])
pressure = int(excode["CH4(sccm)"])
temp = int(excode["温度(℃)"])
sub1 = excode["衬底1"]
sub2 = excode["衬底2"]
metaltype = excode["金属网"]
note = excode["实验目的"]
time = int(excode["持续时间(min)"])
SR = int(excode["方阻(kΩ/□)"])

""" #开始对EXCEL进行编辑
#创建app进程
app = xw.App(visible=False,add_book = False)
# 链接工作表
wb = app.books.open('Report_Quartz_2019_Condition.xlsx')
# 对指定工作表进行编辑
sht = wb.sheets['Ratio MetaData']
# 方式2——显性
# wb = xw.Book('Report_Quartz_2019_Condition.xlsx')
# sht = wb.sheets['Ratio Metadata']
# 获取当前EXCEL表格的行数与列数
info = sht.range('A1').expand('table')
row = info.last_cell.row
col = info.last_cell.column
print(info)
# 计算出要添加的一行位置

rowl =str(row + 1)
print('数据添加所在行：'+rowl)
row = str(row)
print('原表格最后一行：'+rowl) """

# 创建数据列
data = [note,metaltype,Ar,H2,CH4,time,power,pressure,temp,SR]
print(data)
#注入实验条件数据
sht.range('N'+rowl,'W'+rowl).value = data
sht.range('AD'+rowl).value = date
#注入Eecel公式
sht.range('A'+rowl).formula = A
sht.range('J'+rowl).formula = J
sht.range('K'+rowl).formula = K
sht.range('L'+rowl).formula = L
sht.range('X'+rowl).formula = X
sht.range('Y'+rowl).formula = Y
sht.range('Z'+rowl).formula = Z

sht.range('AA'+rowl).formula = AA
sht.range('AB'+rowl).formula = AB
sht.range('AC'+rowl).formula = AC
#保存文件
wb.save()
# 关闭文件
wb.close()
# 结束进程
app.kill