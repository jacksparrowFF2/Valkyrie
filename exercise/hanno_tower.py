#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   tower.py
@Time    :   2020/02/20 20:17:29
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib
# a-c
# a-b
# c-b
# a-c
# b-a
# b-c
# a-c

def move(n,a,b,c):
    if n == 1:
      print(a,"-->",c)
    else:
      move(n-1,a,c,b)
      print(a,"-->",c)
      move(n-1,b,a,c)
      

move(4,'A','B','C')
        
