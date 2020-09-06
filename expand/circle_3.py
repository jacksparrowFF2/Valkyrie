#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   circle_3d.py
@Time    :   2020/09/05 20:34:05
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''
# here put the import lib
import numpy as np
import xlwings as xw
import argparse

parser = argparse.ArgumentParser(description = 'create excel in your select path')

parser.add_argument('-i','--input', metavar='', type=str, required = True, help = 'where your want to creat excel')

args = parser.parse_args()
# hear is the code 
# 函数构建
# 切片次数
t = 900
n = 10
# 生成 z 轴坐标序列
Z = list(range(0,t,1))
# print(Z)
# 生成高度序列
H = list(np.linspace(0.000000000000001,40,t))
# print(H)
# 生成半径序列
R = [x/4 + (360*90)/x for x in H]
# 限定 X 轴坐标
x0 = np.linspace(-450,-360,n)
x1 = np.linspace(-360,-180,4*n)
temp1 = np.linspace(0,180,4*n)
x2= np.linspace(-180,180,8*n)
x3 = np.linspace(180,360,4*n)
temp2 = np.linspace(-180,0,4*n)
x4 = np.linspace(360,450,n)
x = list(x0) + list(x1) + list(x2) + list(x3) + list(x4)
#x = [0,1,2,3,4]
# print(x)
# Z轴循环标致
m = 0
# 设定坐标空集
verts = []
# 计算坐标
for i in range(len(H)):
    z = [Z[m]]*len(x)
    # print(z)
    m = m + 1
    # 通过一般方程计算纵坐标
    y0 = y4 = [H[i]]*len(x0)
    y1 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp1))
    y2 = R[i]-np.sqrt(np.square(R[i])-np.square(x2))
    y3 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp2))
    y = list(y0) + list(y1) + list(y2) + list(y3) + list(y4)
    v = list(zip(x,y,z))
    # 添加坐标点
    verts.extend(v)

# 输出坐标点
# verts = [(0,0,0),(1,0,0),(2,0,0),(3,0,0),(4,0,0),(0,1,0),(1,1,0),(2,1,0),(3,1,0),(4,1,0)]
# print(verts)
print(len(verts))
""" # 构造边 x 方向
print("构造边 x 方向")
edge_xa = list(range(0,len(verts),1))
edge_xb = list(range(1,len(verts)+1,1))
edge_xc = list(zip(edge_xa,edge_xb))
# 构造清除序列
# 乘以横坐标的数量 len(x)
temp_clear_sequence = [1] * len(x)
temp_clear_sequence[-1] = 0
# 乘以切片的数量 t
clear_sequence = temp_clear_sequence * t
# print(clear_sequence)
# 清洗边连接顺序
temp_edges_x = list(map(lambda x,y: x*y,edge_xc,clear_sequence))
# 输出x方向边连接顺序
edges_x = list(filter(None, temp_edges_x))
# print(edges_x)

if t == 1:
    edges_y = []
else:
    print("构造边 y 方向")
    edge_ya = list(range(0,len(verts),1))
    edge_yb = list(range(len(x),len(verts)+len(x),1))
    edge_yc = list(zip(edge_ya,edge_yb))
    edges_y = edge_yc[0:-len(x)]
    print(edges_y)

edges = edges_x + edges_y
# print(edges)

#构建面
print("开始构建面") 
face_a = edges_x
face_b = edges_x[len(x)-1:]
print(face_b)
# 将列表所有元素从tuple转为list
face_b = list(map(list,face_b))
# 将列表所有子list进行翻转
face_b = list(map(reversed,face_b))
# 将翻转后list转换为tuple
face_b = list(map(tuple,face_b))
print(face_b)

# 构建连点顺序
faces = list(map(lambda x,y: x+y,face_a,face_b))
print(faces) """

verts = list(map(str,verts))
verts = [x.replace('(','').replace(')','') for x in verts]

if __name__ == '__main__':
    try:        
        app = xw.App(visible=True,add_book=False)
        # wb = app.books.add()
        wb = app.books.add()
        # wb.sheets.add("sheet2")
        wb.sheets["sheet1"].name = "mesh"
        sht = wb.sheets['mesh']
        name = ["X,Y,Z"]
        print('开始写入坐标')
        # 获取表格坐标信息
        info = sht.range('A1').expand('table')
        row = info.last_cell.row
        col = info.last_cell.column
        print('原表格最后一列：'+str(col))
        # 填写坐标名称
        sht.range('A1').value = name
        # 填写坐标数据
        sht.range((1,col),(len(verts)+1,col)).options(transpose = True).value = verts
        print('注入完成')
    finally:
        if wb:
            wb.save(args.input)
            # wb.close()
            # app.kill()