#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   calculate_gram_needed.py
@Time    :   2019/11/30 20:04:34
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import argparse

# 创建命令解释器
parser = argparse.ArgumentParser('该脚本旨在帮助你计算指定摩尔浓度的溶液所需要的粉末质量')

parser.add_argument('-m', '--mol_mass', metavar = '', type = float, required = True,
                    help = 'g/mol/摩尔每克')
parser.add_argument('-s', '--solution_volume', metavar = '', type = float, required = True,
                    help = 'ml/毫升')
parser.add_argument('-c', '--concentration_coefficient', metavar = '', type = float, required = True,
                    help = '浓度系数')
parser.add_argument('-u', '--unit', metavar = '', type = float, required = True,
                    help = 'mmol/L:毫摩尔每升')
parser.add_argument('-p', '--purity', metavar = '', type = float, required = True,
                    help = '纯度')

args = parser.parse_args()

def calculate_mass(m,s,c,u,p):
    m = args.mol_mass
    s = args.solution_volume
    c = args.concentration_coefficient
    u = args.unit
    p = args.purity
    mass = ((s/1000)*c*u*m/1000)/p
    return mass

if __name__ == '__main__':
    mass = calculate_mass(args.mol_mass,
                          args.solution_volume,
                          args.concentration_coefficient,
                          args.unit,
                          args.purity)
    print(mass)
    print('请输入 -h 以查看使用说明')
    input("Press <enter>")