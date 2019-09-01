#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   add_datapy.py
@Time    :   2019/09/01 20:14:24
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2017-2018, Liugroup-NLPR-CASIA
@Desc    :   None
'''

# here put the import lib
import win32clipboard as wc
import win32con
import ast
import xlwings as xw

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

date = int(excode["日期"])
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

#开始对EXCEL进行编辑
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
print(info)
row = info.last_cell.row
col = info.last_cell.column
# 计算出要添加的一行位置
rowl = row + 1
rowl = str(rowl)
print(rowl)
# 创建数据列
data = [note,metaltype,Ar,H2,CH4,time,power,pressure,temp,SR]
print(data)
#写入最下面一行的数据
sht.range('N'+rowl,'W'+rowl).value = data
sht.range('AE'+rowl).value = date

#保存文件
wb.save()
# 关闭文件
wb.close()
# 结束进程
app.kill