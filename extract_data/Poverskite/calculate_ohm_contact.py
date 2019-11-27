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
import random
V = [1,1,1,1,1,1,0.2,0.1,-0.1,-1,-1,-1]
I = [1,1,1,1,1,1,0.2,0.1,-0.1,-1,-1,-1]


def getposition(inputlist):
    for i in range(len(inputlist)):
        if inputlist[i] > inputlist[i + 1]:
            position_max = i+1
            break
    position_min = inputlist.index(min(inputlist))
    return[position_max,position_min,inputlist[position_max:position_min]]

# def compute_average(V,I,Count):
#     resultList=random.sample(range(pos[0],pos[1]),Count); # sample(x,y)函数的作用是从序列x中，随机选择y个不重复的元素。上面的方法写了那么多，其实Python一句话就完成了。
#     print(resultList)# 打印结果
#     ohm1 = V[resultList[0]]/I[resultList[0]]
#     ohm2 = V[resultList[1]]/I[resultList[1]]
#     ohm3 = V[resultList[2]]/I[resultList[2]]
#     return[ohm1,ohm2,ohm3] 


def compute_average(ohm_input,Count):
    resultList=random.sample(range(pos[0],pos[1]),Count); # sample(x,y)函数的作用是从序列x中，随机选择y个不重复的元素。上面的方法写了那么多，其实Python一句话就完成了。
    print(resultList)# 打印结果
    ohm1 = V[resultList[0]]/I[resultList[0]]
    ohm2 = V[resultList[1]]/I[resultList[1]]
    ohm3 = V[resultList[2]]/I[resultList[2]]
    return[ohm1,ohm2,ohm3] 


#输出取值范围
pos = getposition(V)
# 输出三个随机值计算结果
ohm = compute_average(V,I,3)
print(ohm)
average_ohm = sum(ohm)/len(ohm)
print(average_ohm)

