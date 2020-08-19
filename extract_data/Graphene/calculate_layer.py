#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   calculate_layer.py
@Time    :   2020/01/06 10:11:57
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import math

# 方法一
wG = 1596
pre_n =11/(wG - 1581.6) -1
n = pow(pre_n,1/1.6)

print(n)
# 方法二
