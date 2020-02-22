#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   prod.py
@Time    :   2020/02/21 16:58:54
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib
from functools import reduce

# def prod(x,y):
#     return x*y

# L = [3,5,7,9]
# r = reduce(prod,L)
# print(r)

# 方法2
def prod(L):
    return reduce(lambda x,y:x*y,L)

L = [3,5,7,9]

r = prod(L)
print(r)