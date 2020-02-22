#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   str2float.py
@Time    :   2020/02/21 17:17:50
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib
from functools import reduce

DIGITS = {"0":0,"1":1,"2":2,"3":3,"4":4,"5":5,"6":6,"7":7,"8":8,"9":9,".":".","-":"-"}

def str2folat(s):
    def char2num(s):
        return DIGITS[s]
    def fn(x,y):
        return x*10+y
    n = s.find(".")
    t = list(map(char2num,s))
    p = 10**len(t[n+1:])
    # print(p)
    t.remove(".")
    if s[0] == '-':
        t.remove("-")
        return -1*reduce(fn,t)/p
    else:
        return reduce(fn,t)/p

a = "1332.3"
b = "-1332.3"

# w = str2float(a)


power1 = str2folat(b)
print(power1)
