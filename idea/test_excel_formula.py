#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   Untitled-1
@Time    :   2019/09/02 10:15:15
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2017-2018, Liugroup-NLPR-CASIA
@Desc    :   None
'''

# here put the import lib
# 从另一文件引入变量——方法1
from excel_formula import A,J,K,L,X,Y,Z,AA,AB,AC
print(A)
print(J)
print(K)
print(L)
print(X)
print(Y)
print(Z)
print(AA)
print(AB)
print(AC)
# 从另一文件引入变量——方法2
# import excel_formula
# print(excel_formula.A)

# 测试
from get_table_info import app,wb,sht,info,row,rowl
print(rowl)

input(“Enter the any press to exit” )