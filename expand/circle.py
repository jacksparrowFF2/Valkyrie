#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   circle.py
@Time    :   2020/09/04 09:45:52
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib

# here is the code
# 定义点类
class Point():
    def __init__(self, x, y):
        self.x = x
        self.y = y

# 定义圆心计算函数
def getCircle(p1, p2, p3):
    x21 = p2.x - p1.x
    y21 = p2.y - p1.y
    x32 = p3.x - p2.x
    y32 = p3.y - p2.y
    # three colinear/使用向量检验该三点是否为共线状态
    if (x21 * y32 - x32 * y21 == 0):
        return None
    xy21 = p2.x * p2.x - p1.x * p1.x + p2.y * p2.y - p1.y * p1.y
    xy32 = p3.x * p3.x - p2.x * p2.x + p3.y * p3.y - p2.y * p2.y
    y0 = (x32 * xy21 - x21 * xy32) / 2 * (y21 * x32 - y32 * x21)
    x0 = (xy21 - 2 * y0 * y21) / (2.0 * x21)
    R = ((p1.x - x0) ** 2 + (p1.y - y0) ** 2) ** 0.5
    return x0, y0, R


p1, p2, p3 = Point(0, 0), Point(4.5, 0), Point(2, 2)
# 一次性输出全部信息：横坐标，纵坐标，半径
print(getCircle(p1, p2, p3))
# 输出横坐标
print(getCircle(p1, p2, p3)[0])
# 输出纵坐标
print(getCircle(p1, p2, p3)[1])
# 输出三点共圆半径
print(getCircle(p1, p2, p3)[2])