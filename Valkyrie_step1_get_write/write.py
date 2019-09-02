#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   write.py
@Time    :   2019/09/02 14:28:07
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
import get
from excel_formula import A, J, K, L, X, Y, Z, AA, AB, AC
from get import app, wb, sht, info, row, rowl

# 获取剪贴板内容


def getCopyText():
    wc.OpenClipboard()
    copy_text = wc.GetClipboardData(win32con.CF_UNICODETEXT)
    wc.CloseClipboard
    return copy_text


# 将字符串类型转变为字典类型
excode = ast.literal_eval(getCopyText())

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
data = [note, metaltype, Ar, H2, CH4, time, power, pressure, temp, SR]
# print(data) # 验证数据列是否正确

# 注入实验条件数据
sht.range('N'+rowl, 'W'+rowl).value = data
sht.range('AD'+rowl).value = date

# 注入Eecel公式
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

# 保存文件
wb.save()
# 关闭文件
wb.close()
# 结束进程
app.kill
