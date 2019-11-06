#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   show_hiden_files.py
@Time    :   2019/11/04 14:31:52
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import os
# import subprocess
pwd = os.path.abspath('.')
d =os.listdir(pwd)
c =  os.listdir(r"C:\Users\nuko\Desktop")
print(pwd)
print(d)
print(c)