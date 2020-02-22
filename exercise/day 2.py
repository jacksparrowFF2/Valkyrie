#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   day 2.py
@Time    :   2020/02/19 11:29:40
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib
# 格式化字符串
A = "neko"
B = "plus"
C = "Hello,{1} {0}".format(A,B)
D = F"Hello,{A} {B}"
print(C)

province_you_live = input("where you lived in?\n")

if province_you_live in ("ABB","BAA","AAA"):
  tax = 1
elif province_you_live == "CCC":
  tax = 2
else:
  tax = 3
print(tax)