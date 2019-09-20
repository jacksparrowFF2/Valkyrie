#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   excel_formula.py
@Time    :   2019/09/02 14:29:12
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib

import get
from get import app, info, row, rowl, sht, wb

global A, K, L, M, N, Y, Z, AA, AB, AC, AD

# 测试用变量 开始
# rowl = 10
# rowl = str(rowl)
# 测试用变量 结束

# A
# Excel原始公式
# =B84&CHAR(10)&AC84
# Excel公式定义
A = '=B'+rowl+'&CHAR(10)&AD'+rowl
# print(A)

# K ID/IG
# Excel原始公式
# =H84/D84
# Excel公式定义
K = '=C'+rowl+'/E'+rowl
# print(K)

# L IG'/IG
# Excel原始公式
# =B84/F84
# Excel公式定义
L = '=I'+rowl+'/E'+rowl
# print(L)

# M ID/ID'
# Excel原始公式
# =C84/G84
# Excel公式定义
M = '=C'+rowl+'/G'+rowl
# print(M)

# N 层数
# Excel原始公式
# =IF(88>J84,45/(88-J84),"bulk")
# Excel公式定义
N = '=IF(88>J'+rowl+',45/(88-J'+rowl+'),"bulk")'
# print(J)

# Y 真实氩气
# Excel原始公式
# =Q84*1.415
# Excel公式定义
Y = '=Q'+rowl+'*1'
# print(Y)

# Z 氢气
# Excel原始公式
# =R84*1.01
# Excel公式定义
Z = '=R'+rowl+'*2'
# print(Z)

# AA
# Excel原始公式
# =S84*0.719
# Excel公式定义
AA = '=S'+rowl+'*3'
# print(AA)

# AB 气体流量比
# Excel原始公式
# =Q84&"/"&R84&"/"&S84
# Excel公式定义
AB = '=Q'+rowl+'&"/"&'+'R'+rowl+'&"/"&'+'S'+rowl
# print(AB)

# AC 真实气体流量
# Excel原始公式
# =Y84&"/"&Z84&"/"&AA84
# Excel公式定义
AC = '=Y'+rowl+'&"/"&'+'Z'+rowl+'&"/"&'+'AA'+rowl
# print(AC)

# AD TAG1
# Excel原始公式
# =P2&"/"&Q2&"/"&R2&"/"&S2&"/"&T2&"/"&U2&"/"&V2&"/"&W2
# Excel公式定义
AD = '=P'+rowl+'&"/"&'+'Q'+rowl+'&"/"&'+'R'+rowl+'&"/"&'+'S'+rowl+'&"/"&'+'T'+rowl+'&"/"&'+'U'+rowl+'&"/"&'+'V'+rowl+'&"/"&'+'W'+rowl
# print(AD)

# 检查公式是否正确，如果正确请注释

# print('A'+A)
# print('K'+K)
# print('L'+L)
# print('M'+M)
# print('N'+N)
# print('Y'+Y)
# print('Z'+Z)
# print('AA'+AA)
# print('AB'+AB)
# print('AC'+AC)
# print('AD'+AD)

# a = sht.range('A'+row).value
# print(a)

""" #保存文件
wb.save()
# 关闭文件
wb.close()
# 结束进程
app.kill """