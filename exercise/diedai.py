#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   diedai.py
@Time    :   2020/02/21 11:28:41
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib

# function
def find_min_and_max(L):
    if L == []:
        return (None,None)
    else:
        ma = L[0]
        mi = L[0]
        for i in L:
            if i > ma:
                ma = i
            if i < mi:
                mi = i
        print(ma)
        print(mi)
    

        
a = [1,3,5,0,0,7,2]
find_min_and_max(a)