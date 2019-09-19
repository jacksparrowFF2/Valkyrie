#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   test7.py
@Time    :   2019/09/19 09:38:43
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
# =INDIRECT("'Ratio Metadata'!$A"&COLUMN())


# a = '=INDIRECT('
# print(a)
# b = '"'
# print(b)
# c = '\'Ratio Metadata\''
# print(c)
# d = '!$A'
# print(d)
# e = '"&COLUMN())'
# print(e)

# formula = a + b + c + d + e
# print(formula)
row = 1
rowl = row + 1
print('原表格最后一行：'+str(row))

# =B2&CHAR(10)&AG2

A ='=B%s&CHAR(10)&AG%s'%(rowl,rowl)
K ='=C%s/E%s'%(rowl,rowl)
L ='=I%s/E%s'%(rowl,rowl)
M ='=C%s/G%s'%(rowl,rowl)
N ='=IF(88>J%s,45/(88-J%s),"bulk")'%(rowl,rowl)
Y = '=Q%s*1.415'%(rowl)
Z = '=R%s*1.01'%(rowl)
AA = '=S%s*0.719'%(rowl)
AB = '=Q%s&"/"&R%s&"/"&S%s'%(rowl,rowl,rowl)
AC = '=Y%s/SUM(Y%s+Z%s+AA%s)'%(rowl,rowl,rowl,rowl)
AD = '=Z%s/SUM(Y%s+Z%s+AA%s)'%(rowl,rowl,rowl,rowl)
AE = '=AA%s/SUM(Y%s+Z%s+AA%s)'%(rowl,rowl,rowl,rowl)
AF = '=P%s&"/"&Q%s&"/"&R%s&"/"&S%s&"/"&T%s&"/"&U%s&"/"&V%s&"/"&W%s'%(rowl,rowl,rowl,rowl,rowl,rowl,rowl,rowl)
print(AA, AB, AC, AD, AE, A, K, L, M, N, Y, Z)
