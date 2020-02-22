#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   huishu.py
@Time    :   2020/02/21 22:17:12
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib

# 方法一
def check(n):
    if n < 10:
        return n
    elif n < 100:
        a = n // 10
        b = n % 10
        if n == b*10+a:
            return n
    elif n < 1000:
        a = n // 100
        b = (n//10) % 10
        c = n % 10
        if n == 100*c+10*b+a:
            return n
    else:
        print('超过1000')
# 方法二
def check2(s):
    sstr = str(s)
    rsstr = sstr[::-1]
    if sstr == rsstr:
        return s


a = check(10)
b = range(1,1000)
print(b)

c = filter(check,b)
d = list(c)
print(d)

e = filter(check2,b)
f = list(e)
print(f)