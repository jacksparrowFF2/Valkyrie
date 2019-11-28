#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   calculate_ohm_contact.py
@Time    :   2019/11/27 20:20:42
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
# 导入随机数组件
import random
# 导入excel操作组件
import xlwings as xw
# 导入命令解释器
import argparse
# 导入 CSV 组件
import csv

def getposition(inputlist):
    for i in range(len(inputlist)):
        if inputlist[i] > inputlist[i + 1]:
            position_max = i+1
            break
    position_min = inputlist.index(min(inputlist))
    return[position_max,position_min,inputlist[position_max:position_min]]

def compute_average_3(V,I,Count):
    resultList=random.sample(range(pos[0],pos[1]),Count); # sample(x,y)函数的作用是从序列x中，随机选择y个不重复的元素。上面的方法写了那么多，其实Python一句话就完成了。
    print(resultList)# 打印结果
    ohm1 = V[resultList[0]]/I[resultList[0]]
    ohm2 = V[resultList[1]]/I[resultList[1]]
    ohm3 = V[resultList[2]]/I[resultList[2]]
    return[ohm1,ohm2,ohm3] 


def compute_average_2(ohm_input,Count):
    resultList=random.sample(range(pos[0],pos[1]),Count); # sample(x,y)函数的作用是从序列x中，随机选择y个不重复的元素。上面的方法写了那么多，其实Python一句话就完成了。
    print(resultList)# 打印结果
    ohm1 = float(R[resultList[0]])
    ohm2 = float(R[resultList[1]])
    ohm3 = float(R[resultList[2]])
    
    return[ohm1,ohm2,ohm3] 

# 创建命令解释器
parser = argparse.ArgumentParser('计算欧姆接触电阻')

parser.add_argument('-i','--input',metavar = '', type = argparse.FileType(mode = 'r'), 
                    help = '要进行处理的 excel 文件')
parser.add_argument('-e','--excel',metavar = '', type = str,
                    help = '读取数据的 excel 文件路径')

group = parser.add_argument_group('进阶选项')
group.add_argument('-create', '--create_excel', action = 'store_true', 
                   help = '模式：将实验条件写入指定 excel 表格')
group.add_argument('-save', '--savedata', action = 'store_true', 
                   help = '模式：将数据写入指定 excel 表格')

args = parser.parse_args()

if __name__ == '__main__':
    if args.create_excel:
        print('a')
    elif args.savedata:
        csvFile = args.input
        reader = csv.reader(csvFile)
        R = []
        I = []
        for item in reader:
            if item[4] == '':
                pass
            else:
                I.append(item[2])
                R.append(item[4])
        I.pop(0)
        R.pop(0)
        # 输出电流
        print(I)
        # 输出电阻
        print(R)
        # 输出取值范围
        pos = getposition(I)
        # 输出三个随机值计算结果
        ohm = compute_average_2(R,3)
        print(ohm)
        average_ohm = sum(ohm)/len(ohm)
        print(average_ohm)
    else:
        print('请输入 -h 以查看使用说明')
        input("Press <enter>")
