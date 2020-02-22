#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   triangles.py
@Time    :   2020/02/21 14:36:37
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''


def triangles():
    n = 1
    a = [1]
    b = [1, 1]
    yield a
    while True:
        a = b
        b = [1, 1]
        for x in list(range(1,n)):
            b.insert(x,a(x-1)+a(x))
            yield b
        n = n + 1


def move(n,a,b,c):
    if n == 1:
        print(a,"-->",c)
    else:
        move(n-1,a,c,b)
        print(a,"-->",c)
        move(n-1,c,a,b)
            
move(3,"A","B","C")