#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   format
@Time    :   2019/09/07 20:34:54
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import win32clipboard as wc
import win32con
import ast
import re
# 正则匹配表达
""" ([\S]+&)
(@[\S]+)
(@[\S]+ +\S+)

"@+(\S+)+(",)+( )+(\S+)+ """

# a = {
#     日期: 20190830_01_01
#     实验目的: 改变
#     实验过程: 没有出现故障
#     初始输入功率(w): 81
#     初始反馈功率(w): 31
#     末端输入功率(w): 86
#     末端反馈功率(w): 36
#     Ar(sccm): 150
#     H2(sccm): 0
#     CH4(sccm): 9
#     压强(pa): 200
#     温度(℃): 600
#     持续时间(min): 60
#     衬底1: n-Si
#     衬底2: Quartz
#     金属网: MK1 铜网0.5_5.0
#     初步实验结果: ssAA啊啊啊
#     方阻(kΩ /□): 1
# }


def getCopyText():
    wc.OpenClipboard()
    copy_text = wc.GetClipboardData(win32con.CF_UNICODETEXT)
    wc.CloseClipboard
    return copy_text


""" # 将字符串类型转变为字典类型
excode = ast.literal_eval(getCopyText())

# 输出变量类型，确保为字典类型
print(type(excode)) """

a = getCopyText()
print(a)
print(type(a))

b = a.split('\n')
print(b)
print(type(b))

c = []
for i in b:
    i = i.replace('\r', '')
    c.append(i.split(':'))
print(c)

d = {}
for i in range(len(c)):
    d[c[i][0]] = c[i][1]
print(d)

print(d["日期"])