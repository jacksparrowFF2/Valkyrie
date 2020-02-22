#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   string_trim.py
@Time    :   2020/02/20 21:04:22
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib
def trim(s):
    if len(s) != 0:
        i = 0
        j = len(s)-1
        t = 0
        while s[i] == " ":
            i = i+1
            t = i
        m = len(s)
        while s[j] == " ":
            m = j
            j = j -1
        print(s[t:m])
    else:
        print(s)

def trim2(s):
    while s!='' and s[0]==" ":
        s=s[1:]
    while s!='' and s[-1]==" ":
        s=s[:-1] 
    print(s)

s = " a b "
# trim(s)
trim2(s)
