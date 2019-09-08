#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   write.py
@Time    :   2019/09/08 12:00:16
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib


# 导入 ast,用于将字符串类型转变为字典类型
import ast
# 导入剪贴板相关模块
import win32clipboard as wc
import win32con
import xlwings as xw
# 导入第一个程序
import get
# 导入变量
from excel_formula import AA, AB, AC, AD, A, K, L, M, N, Y, Z
from get import app, info, row, rowl, sht, wb

# 获取剪贴板内容
def getCopyText():
    wc.OpenClipboard()
    copy_text = wc.GetClipboardData(win32con.CF_UNICODETEXT)
    wc.CloseClipboard
    return copy_text


# 开始对剪贴板内容进行格式化，格式化为字典
excode_a = getCopyText()
excode_b = excode_a.split('\n')
excode_c = []
for i in excode_b:
    i = i.replace('\r', '')
    excode_c.append(i.split(':'))

excode = {}
for i in range(len(excode_c)):
    excode[excode_c[i][0]] = excode_c[i][1]
print(excode)

# 输出变量类型，确保为字典类型
print(type(excode))

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
SR = int(excode["方阻(kΩ/□)"])

# 创建实验条件数据列
data = [note+"+"+sub1+"+"+sub2, metaltype, Ar, H2, CH4, time, power, pressure, temp, SR]
# print(data) # 验证数据列是否正确

# 注入实验条件数据
sht.range('O'+rowl, 'X'+rowl).value = data
sht.range('B'+rowl).value = date

# 注入Eecel公式
sht.range('A'+rowl).formula = A
sht.range('K'+rowl).formula = K
sht.range('L'+rowl).formula = L
sht.range('M'+rowl).formula = M
sht.range('N'+rowl).formula = N
sht.range('Y'+rowl).formula = Y
sht.range('Z'+rowl).formula = Z
sht.range('AA'+rowl).formula = AA
sht.range('AB'+rowl).formula = AB
sht.range('AC'+rowl).formula = AC
sht.range('AD'+rowl).formula = AD

# 保存文件
wb.save()
# 关闭文件
wb.close()
# 结束进程
app.kill
