#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   dark_IV.py
@Time    :   2019/12/28 11:00:37
@Author  :   SPH
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import argparse
import win32clipboard as w
import win32con
import numpy as np
import math

parser = argparse.ArgumentParser(description='该脚本用于对暗态IV数据进行处理')

parser.add_argument('-i', '--input', metavar='',
                    type=argparse.FileType(mode='r'))
parser.add_argument('-n', '--factor', metavar='',
                    type=float)

parser.add_argument_group('基础选项')
parser.add_argument('-c', '--copy', action='store_true', help='将数据写入到剪贴板中')
parser.add_argument('-s1', '--step1', action='store_true', help='第一阶段数据处理')
parser.add_argument('-s2', '--step2', action='store_true', help='第二阶段数据处理')

groupA = parser.add_mutually_exclusive_group()
# parser.add_mutually_exclusive_group('高级选项')
groupA.add_argument('-l', '--light', action='store_true', help='光照IV')
groupA.add_argument('-d', '--dark', action='store_true', help='暗态IV')

groupB = parser.add_mutually_exclusive_group()
# parser.add_mutually_exclusive_group('高级选项')
groupB.add_argument('-f1', '--floor1', action='store_true', help='一楼IV测试设备')
groupB.add_argument('-f2', '--floor2', action='store_true', help='二楼IV测试设备')

args = parser.parse_args()


# Richard Constant(A cm^-2 k^-2)
RC = 252
# temperature(℃)
t = 25
# absolute temperature(k)
T = t + 273.15
# reverse bias saturation current density
J0 = 1
# Boltzmann Constant(J/k)
kB = 1.380649E-23
# electron 
elec = 1.6021766208E-19
# 构建复合因子
beta = elec/(kB*T)

def writeclip(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, astring)
    w.CloseClipboard()
def data_process(parameter_list):
    temp_list = []
    for i in parameter_list:
        x = '\t'.join(str(num) for num in i)
        # print(x)
        temp_list.append(x)
    print('输出第二阶段数据处理结果')
    # print(temp_list)
    out_list = []
    for i in temp_list:
        x = i + '\n'
        out_list.append(x)
    print('输出第三阶段数据处理结果')
    # print(out_list)
    out_str = ''.join(out_list)
    print('输出第三阶段数据处理结果')
    return out_str

if __name__ == '__main__':
    if args.floor1:
        if args.dark:
            if args.copy:
                infile = args.input
                All_data = infile.readlines()
                # print(All_data)
                Area = float(All_data[7].split('\t')[1].split('\n')[0])
                V = []
                I = []
                J = []
                for line in All_data[13:]:
                    if line != '\n':
                        V1,I1 = line.split('\t')
                        # print(V1,I1)
                        V.append(float(V1))
                        I.append(float(I1))
                # print(V)
                # print(I)
                # 构建 J 数列
                J = list(map(lambda x : x/Area, I))
                # 构建 1/(J*1000)
                J_1000 = list(map(lambda x: 1/(x*1000), J))
                # 构建对数处理 J 数列
                absJ = list(map(abs, J))
                # 进行对数处理
                lnJ = list(map(math.log, absJ))
                
                if args.step1:
                    name = ['V', 'I', 'J', '1/J_1000', 'absJ', 'lnJ']
                    Unit = ['V', 'A', 'A/cm2', 'A-1cm-2', 'A/cm2', 'A/cm2']
                    temp_list1 = [list(item) for item in zip(V, I, J, J_1000, absJ,  lnJ)]
                    temp_list1.insert(0,'')
                    temp_list1.insert(0,Unit)
                    temp_list1.insert(0,name)
                    # print(temp_list1)
                    out_str = data_process(temp_list1)
                    # print(out_str)
                    writeclip(out_str)
                elif args.step2:
                    # 创建HJ数列
                    HJ_temp = list(map(lambda x: args.factor*math.log(x/(RC*T**2))/beta, J))
                    HJ = list(np.array(V) - np.array(HJ_temp))
                    name = ['HJ']
                    Unit = ['A/cm2']
                    temp_list1 = [list(item) for item in zip(HJ)]
                    # print(temp_list1)
                    temp_list1.insert(0,'')
                    temp_list1.insert(0,Unit)
                    temp_list1.insert(0,name)
                    out_str = data_process(temp_list1)
                    # print(out_str)
                    writeclip(out_str)
                else:
                    print('请选择处理阶段')
        print('处理结束')
    elif args.floor2:
        if args.dark:
            if args.copy:
                infile = args.input
                All_data = infile.readlines()
                # print(All_data)
                index = []
                V = []
                I = []
                Power = []
                J = []
                for line in All_data[1:]:
                    index1, V1, I1, Power1, J1 = line.split()
                    index.append(float(index1))
                    V.append(float(V1))
                    I.append(float(I1))
                    Power.append(float(Power1))
                    J.append(float(J1))
                    # index= line.split()
                # 构建对数处理 I 数列
                I = list(map(abs, I))
                # 进行对数处理
                lnI = list(map(math.log, I))
                Power = list(map(abs, Power))
                # 构建对数处理 J 数列
                J = list(map(abs, J))
                # 进行对数处理
                lnJ = list(map(math.log, J))
                J_1000 = list(map(lambda x: x*1000, J))
                # 构建正常 J 数组，单位为 mA/cm2
                rJ_1000 = list(map(lambda x: -x*1000, J))
                
                # print(HJ)
                if args.step1:
                    name = ['Voltage', 'I', 'Power', 'J', 'lnJ', 'J_1000', 'rJ_1000']
                    Unit = ['V', 'A', 'w', 'A/cm2', 'A/cm2', 'mA/cm2', 'mA/cm2']
                    temp_list1 = [list(item) for item in zip(V, I, Power, J, lnJ, J_1000, rJ_1000)]
                    temp_list1.insert(0,'')
                    temp_list1.insert(0,Unit)
                    temp_list1.insert(0,name)
                    # print(temp_list1)
                    out_str = data_process(temp_list1)
                    print(out_str)
                    writeclip(out_str)

                    # temp_list2 = []
                    # for i in temp_list1:
                    #     x = '\t'.join(str(num) for num in i)
                    #     # print(x)
                    #     temp_list2.append(x)
                    # print('输出第二阶段数据处理结果')
                    # # print(temp_list2)
                    # out_list = []
                    # for i in temp_list2:
                    #     x = i + '\n'
                    #     out_list.append(x)
                    # print('输出第三阶段数据处理结果')
                    # # print(out_list)
                    # out_str = ''.join(out_list)
                    # print('输出第三阶段数据处理结果')
                    # # print(out_str)
                    # writeclip(out_str)
                elif args.step2:
                    # 创建HJ数列
                    HJ_temp = list(map(lambda x: args.factor*math.log(x/(RC*T**2))/beta, J))
                    HJ = list(np.array(V) - np.array(HJ_temp))
                    name = ['HJ']
                    Unit = ['A/cm2']
                    temp_list1 = [list(item) for item in zip(HJ)]
                    # print(temp_list1)
                    temp_list1.insert(0,'')
                    temp_list1.insert(0,Unit)
                    temp_list1.insert(0,name)
                    out_str = data_process(temp_list1)
                    # print(out_str)
                    writeclip(out_str)
                else:
                    print('请选择处理阶段')
        print('处理结束')
    else:
      print('请选择测试设备所在楼层、数据处理阶段')
