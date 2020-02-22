#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   liebiaoshengchengshi.py
@Time    :   2020/02/21 13:06:47
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib
L = ['Hello', 'World', 18, 'Apple', None]
# L = ['Hello', 'World', 'Apple']
a = [s.lower() for s in L if isinstance(s,str)]
# a = [s.lower() for s in L]
print(a)