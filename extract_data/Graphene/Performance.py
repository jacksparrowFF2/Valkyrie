#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   Performance.py
@Time    :   2019/12/16 20:18:53
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import math

# Richard Constant(A cm^-2 k^-2)
A = 252
# temperature(℃)
t = 25
# absolute temperature(k)
T = t + 273.15
print(T)
# reverse bias saturation current density
J0 = 1
# Boltzmann Constant(J/k)
kB = 1.380649E-23
# electron 
elec = 1.6021766208E-19
# 构建复合因子
a = elec/(kB*T)
a2 = (kB*T)/elec
print(a)
print(a2)


# 引入J-V数据（A cm^-2-V）
V = 0.25
J = 0.25
# 计算斜率（无量纲量）
slope = 23.55





# 计算理想因子/品质因子
n1 = a/slope
print(n1)
# 计算反向饱和电流/暗态饱和电流密度
# phi = math.log(10)

# print(phi)