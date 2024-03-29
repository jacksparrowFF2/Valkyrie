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

# here put the import lib
import bpy
import numpy as np
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
# 起始直线段的X坐标
x00 = np.linspace(-550,-450,n)
# 第一段圆的X坐标
x0 = np.linspace(-450,-270,n)
temp0 = np.linspace(0,180,n)
# print(x0)
# 第二段圆的X坐标
x1 = np.linspace(-270,-90,n)
temp1 = np.linspace(-180,0,n)
# 第三段直线的坐标
x2 = np.linspace(-90,90,n)
# 第三段圆的X坐标
x3 = np.linspace(90,270,n)
temp3 = np.linspace(0,180,n)
# 第四段圆的X坐标
x4 = np.linspace(270,450,n)
temp4 = np.linspace(-180,0,n)
# 终结直线段的X坐标
x01 = np.linspace(450,550,n)
# 总X轴坐标
x = list(x00)+list(x0) + list(x1) + list(x2) + list(x3) + list(x4) + list(x01)
# x = list(x0) + list(x1) + list(x2)
# # print(x)
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
    y01 = [40]*len(x00)
    y0 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp0))+float(40-H[i])
    y1 = R[i]-np.sqrt(np.square(R[i])-np.square(temp1))+float(40-H[i])
    y2 = [40-H[i]]*len(x2)
    y3 = R[i]-np.sqrt(np.square(R[i])-np.square(temp3))+float(40-H[i])
    y4 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp4))+float(40-H[i])
    y02 = [40]*len(x01)
    # y3 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp2))
    y = list(y01) + list(y0) + list(y1) + list(y2) + list(y3) + list(y4) + list(y02)
    v = list(zip(x,y,z))
    # 添加坐标点
    verts.extend(v)

# 输出坐标点
#verts = [(0,0,0),(1,0,0),(2,0,0),(3,0,0),(4,0,0),(0,1,0),(1,1,0),(2,1,0),(3,1,0),(4,1,0)]
print(verts)
print(len(verts))
print(len(x))
# 构造边 x 方向
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
print(edges_x)

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
print(edges)

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
print(faces)

# 新建网格
mesh = bpy.data.meshes.new('Pyramid_mesh')
# 载入网格数据
mesh.from_pydata(verts,edges,faces)
mesh.update()
# 新建物体”SSS“，并使用”mesh“网格数据
sss = bpy.data.objects.new('SSS',mesh)
scene = bpy.context.scene
# 将物体连接至场景
bpy.context.scene.collection.objects.link(sss)

