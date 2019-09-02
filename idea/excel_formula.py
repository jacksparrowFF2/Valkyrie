#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   excel_formula.py
@Time    :   2019/09/02 11:08:28
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, Liugroup-NLPR-CASIA
@Desc    :   None
'''

# here put the import lib
import get_table_info
global A,J,K,L,X,Y,Z,AA,AB,AC
from get_table_info import app,wb,sht,info,row,rowl

# 测试用变量 开始
# rowl = 10
# rowl = str(rowl)
# 测试用变量 结束

# A
# Excel原始公式
# =AD84&CHAR(10)&AC84
# Excel公式定义
A = '=AD'+rowl+'&CHAR(10)&AC'+rowl
# print(A)

# J
# Excel原始公式
# =B84/D84
# Excel公式定义
J = '=B'+rowl+'/D'+rowl
# print(J)

# K
# Excel原始公式
# =H84/D84
# Excel公式定义
K = '=H'+rowl+'/D'+rowl
# print(K)

# L
# Excel原始公式
# =B84/F84
# Excel公式定义
L = '=B'+rowl+'/F'+rowl
# print(L)

# M
# Excel原始公式
# =IF(88>I84,45/(88-I84),"bulk")
# Excel公式定义
M = '=IF(88>I'+rowl+',45/(88-I'+rowl+'),"bulk")'
# print(M)

# X
# Excel原始公式
# =P84*1.415
# Excel公式定义
X = '=P'+rowl+'*1.415'
# print(X)

# Y
# Excel原始公式
# =Q84*1.01
# Excel公式定义
Y = '=Q'+rowl+'*1.01'
# print(Y)

# Z
# Excel原始公式
# =R84*0.719
# Excel公式定义
Z = '=R'+rowl+'*0.719'
# print(Z)

# AA
# Excel原始公式
# =P84&"/"&Q84&"/"&R84
# Excel公式定义
AA = '=P'+rowl+'&"/"&'+'Q'+rowl+'&"/"&'+'R'+rowl
# print(AA)

# AB
# Excel原始公式
# =X84&"/"&Y84&"/"&Z84
# Excel公式定义
AB  = '=X'+rowl+'&"/"&'+'Y'+rowl+'&"/"&'+'Z'+rowl
# print(AB)

# TAG1
# Excel原始公式
# =O84&"/"&P84&"/"&Q84&"/"&R84&"/"&S84&"/"&T84&"/"&U84&"/"&V84
# Excel公式定义
AC = '=O'+rowl+'&"/"&'+'P'+rowl+'&"/"&'+'Q'+rowl+'&"/"&'+'R'+rowl+'&"/"&'+'S'+rowl+'&"/"&'+'T'+rowl+'&"/"&'+'U'+rowl+'&"/"&'+'V'+rowl
# print(TAG1)




