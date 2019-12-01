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
import pandas as pd
import numpy as np
from pandas import Series,DataFrame
from numpy import nan as NaN

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

parser.add_argument('-i','--input',metavar = '', type = str, 
                    help = '要进行处理的 csv 文件')
parser.add_argument('-e','--excel',metavar = '', type = str,
                    help = '整合数据的 excel 文件路径')

group = parser.add_argument_group('进阶选项')
group.add_argument('-creat', '--creat_excel', action = 'store_true', 
                   help = '模式：将实验条件写入指定 excel 表格')
group.add_argument('-save', '--savedata', action = 'store_true', 
                   help = '模式：计算该电池的平均电阻并写入指定 excel 表格')
group.add_argument('-average', '--average', action = 'store_true', 
                   help = '模式：计算该系列的最终平均电阻并写入指定 excel 表格')
group.add_argument('-p', '--performance', action = 'store_true', 
                   help = '模式：计算该系列的最终平均电阻并写入指定 excel 表格')

args = parser.parse_args()

if __name__ == '__main__':
    if args.creat_excel:
        try:
            app = xw.App(visible = True, add_book = False)
            wb = app.books.add()
            wb.sheets["sheet1"].name = "Average ohm"
            sht = wb.sheets['Average ohm']
            name = ["ohm1", "ohm2", "ohm3", "Average ohm"]
            sht.range('A1','D1').value = name
            # 格式化表格
            # 对第一行标题进行格式化
            sht.range('A1').expand('right').api.HorizontalAlignment = -4108
            sht.range('A1').expand('right').api.VerticalAlignment = -4108
            # 行高
            sht.api.Rows(1).RowHeight = 20
            # 列宽
            sht.api.Columns("A:D").Columnwidth = 15
        finally:
            if wb:
              wb.save(args.excel)
              wb.close()
              app.kill()
    elif args.savedata:
        csvFile = args.input
        # 输出电流数据列
        I = pd.read_csv(csvFile,skiprows=[0,1,2,3,4,5,6,7,8,9],engine='python',usecols = [2])
        R = pd.read_csv(csvFile,skiprows=[0,1,2,3,4,5,6,7,8,9],engine='python',usecols = [4])
        # 将输出的电流数据列转化为 list 
        I = I['I(mA)'].tolist()
        R = R['R(ohm)'].tolist()
        # # 输出电流
        # print(I)
        # # 输出电阻
        # print(R)
        
        # 删除最后一个空元素
        I.pop(len(I)-1)
        # 对电流取整并转换至列表
        I = list(map(round,I))
        # 输出取值范围
        pos = getposition(I)
        print(pos)
        # 输出三个随机值计算结果
        ohm = compute_average_2(R,3)
        print(ohm)
        average_ohm = sum(ohm)/len(ohm)
        print(average_ohm)
        ohm.append(average_ohm)
        abs_ohm = list(map(abs,ohm))
        print(abs_ohm)
        
        
        # 写入ohm
        try:
            inexcel = args.excel
            app = xw.App(visible = False, add_book = False)
            wb = app.books.open(inexcel)
            sht = wb.sheets['Average ohm']
            info = sht.range('A1').expand('table')
            row = info.last_cell.row
            column = info.last_cell.column
            rowl = row + 1
            sht.range('A'+str(rowl),'D'+str(rowl)).value = abs_ohm
            # 格式化表格
            # 对第一行标题进行格式化
            sht.range('A'+str(rowl)).expand('right').api.HorizontalAlignment = -4108
            sht.range('A'+str(rowl)).expand('right').api.VerticalAlignment = -4108
            # 行高
            sht.api.Rows(rowl).RowHeight = 20
        finally:
            if wb:
              wb.save()
              wb.close()
              app.kill()
    elif args.performance:
        csvFile = args.input
        data = DataFrame([[12,'man','13865626962'],[19,'woman',NaN],[17,NaN,NaN],[NaN,NaN,NaN]],columns=['age','sex','phone'])
        print(type(data))
        a = data.dropna(axis=0,how="any")
        print(a)
        # 输出电流数据列
        I = pd.read_csv(csvFile,skiprows=[0,1,2,3,4,5,6,7,8,9] ,engine='python',usecols = [6,7,8,9,10,11,12,13])
        # 将输出的电流数据列转化为 list 
        # I = I['I(mA)'].tolist()
        # 输出电流
        # print(I)
        print(type(I))
        # I.dropna(axis=0, how='all')
        # I.fillna(0)
        # print(I)
        # performance = np.array(I).tolist()
        # a = performance[0]
        # print(a)
        # for item in performance:
        #     if item[0] == a:
        #         pass
        #     else:
        #         print(item)
         
        
        # print()
        # print(performance[-1])
        # a = []
        # for item in performance[-1]:
        #     item = float(item)
        #     a.append(item)
            
        # print(a)
        # # 写入ohm
        # try:
        #     inexcel = args.excel
        #     app = xw.App(visible = False, add_book = False)
        #     wb = app.books.open(inexcel)
        #     sht = wb.sheets['Average ohm']
        #     info = sht.range('A1').expand('table')
        #     row = info.last_cell.row
        #     column = info.last_cell.column
        #     rowl = row + 1
        #     sht.range('A'+str(rowl),'D'+str(rowl)).value = abs_ohm
        #     # 格式化表格
        #     # 对第一行标题进行格式化
        #     sht.range('A'+str(rowl)).expand('right').api.HorizontalAlignment = -4108
        #     sht.range('A'+str(rowl)).expand('right').api.VerticalAlignment = -4108
        #     # 行高
        #     sht.api.Rows(rowl).RowHeight = 20
        # finally:
        #     if wb:
        #       wb.save()
        #       wb.close()
        #       app.kill()
    elif args.average:
        # 写入ohm
        try:
            inexcel = args.excel
            app = xw.App(visible = False, add_book = False)
            wb = app.books.open(inexcel)
            sht = wb.sheets['Average ohm']
            info = sht.range('A1').expand('table')
            row = info.last_cell.row
            column = info.last_cell.column
            rowl = row + 1
            sht.range('D'+str(rowl)).formula = '=sum(D2:D%s)/(%s-2)'%(row,rowl)
            # 格式化表格
            # 对第一行标题进行格式化
            sht.range('D'+str(rowl)).api.HorizontalAlignment = -4108
            sht.range('D'+str(rowl)).api.VerticalAlignment = -4108
            # 行高
            sht.api.Rows(rowl).RowHeight = 20
        finally:
            if wb:
              wb.save()
              wb.close()
              app.kill()
    else:
        print('请输入 -h 以查看使用说明')
        input("Press <enter>")

        # reader = csv.reader(csvFile)
        # for item in reader:
        #     print(item[4])
        # R = []
        # I = []
        # for item in reader:
        #     if item[4] == '':
        #         pass
        #     else:
        #         I.append(item[2])
        #         R.append(item[4])