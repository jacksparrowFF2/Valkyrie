#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   circle2.py
@Time    :   2020/09/04 14:35:21
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import matplotlib.pyplot as plt
# hear is the code 
# 获取用户输入高度H值
H = float(input("please enter H value: "))
print(H)
# 圆弧半径R
R =  H/4 + (360*90)/H
print(R)

# 画圆
fig = plt.figure(figsize=(200, 200))
# 第一个圆
circle = plt.Circle((0, R), R, color='y', fill=False)
plt.gcf().gca().add_artist(circle)
# 第二个圆
circle = plt.Circle((360, H-R), R, color='r', fill=False)
plt.gcf().gca().add_artist(circle)
# plt.axis('equal')
plt.xlim(0,400)
plt.ylim(0,100)
plt.show()