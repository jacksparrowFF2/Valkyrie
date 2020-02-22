#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   normalize_name.py
@Time    :   2020/02/21 16:37:58
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib
def normalize(s):
    s = s.lower()
    s = s[0].upper() + s[1:]
    return s

L1 = ['adam', 'LISA', 'barT']
print(L1)
L2 = list(map(normalize,L1))
print(L2)