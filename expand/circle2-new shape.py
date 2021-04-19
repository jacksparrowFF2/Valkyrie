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
import numpy as np
import xlwings as xw
import argparse
# hear is the code 
 # 函数构建
# 切片次数=细分工件高度次数
t = 20
# 精度？
n = 180
# 生成 z 轴坐标序列（层数）
Z = list(range(0,t,1))
# print(Z)
# 生成高度序列=每层的高度
H = list(np.linspace(0.04,40,t))
# H = [0.04,2.18,2.22,4.44,6.67,8.89,11.11,13.33,15.56,17.78,20.00,22.22,24.44,26.67,28.89,31.11,33.33,35.56,37.78,40.00]
print(H)
# 生成半径序列
R = [x/4 + (360*90)/x for x in H]
print(R)

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
# Z轴循环标志
m = 0
for i in range(len(H)):
    z = [Z[m]]*len(x)
    # print(z)
    m = m + 1
    # 一般方程
    y01 = [40]*len(x00)
    y0 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp0))+float(40-H[i])
    y1 = R[i]-np.sqrt(np.square(R[i])-np.square(temp1))+float(40-H[i])
    y2 = [40-H[i]]*len(x2)
    y3 = R[i]-np.sqrt(np.square(R[i])-np.square(temp3))+float(40-H[i])
    y4 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp4))+float(40-H[i])
    y02 = [40]*len(x01)
    # y3 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp2))
    y = list(y01) + list(y0) + list(y1) + list(y2) + list(y3) + list(y4) + list(y02)
    # y = list(y0) + list(y1) + list(y2)
    # print(y)
    # 画图
    # plt.axis('equal')
    # plt.ylim(0, 50)
    plt.plot(x,y)
    plt.pause(0.05)
plt.show() 

""" 
# 创建 Excel
parser = argparse.ArgumentParser(description = 'create excel in your select path')

parser.add_argument('-i','--input', metavar='', type=str, required = True, help = 'where your want to creat excel')

args = parser.parse_args()

print("开始创建excel")
if __name__ == '__main__':
    try:        
        app = xw.App(visible=True,add_book=False)
        # wb = app.books.add()
        wb = app.books.add()
        # wb.sheets.add("sheet2")
        wb.sheets["sheet1"].name = "circle"
        sht = wb.sheets['circle']
        name = ["X", "Y", "Z"]
        
        print('开始计算数据并写入')
        # 函数构建
        # 切片次数
        t = 900
        n = 90
        # 生成 z 轴坐标序列
        Z = list(reversed(range(0, t, 1)))
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
        # print(x)
        m = 0
        for i in range(len(H)):
            z = [Z[m]]*len(x)
            # print(z)
            m = m + 1
            # 一般方程
            y0 = y4 = [H[i]]*len(x0)
            y1 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp1))
            y2 = R[i]-np.sqrt(np.square(R[i])-np.square(x2))
            y3 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp2))
            y = list(y0) + list(y1) + list(y2) + list(y3) + list(y4)
            # print(y)
            # 获取表格坐标信息
            info = sht.range('A1').expand('table')
            row = info.last_cell.row
            col = info.last_cell.column
            print('原表格最后一列：'+str(col))
            if col == 1:
                # 填写坐标名称
                sht.range((1,col),(1,col+2)).options(transpose = False).value = name
                # 填写X坐标数据
                sht.range((2,col),(2+len(x),col)).options(transpose = True).value = x
                # 填写y坐标数据
                sht.range((2,col+1),(2+len(x),col+1)).options(transpose = True).value = y
                # 填写z坐标数据
                sht.range((2,col+2),(2+len(x),col+2)).options(transpose = True).value = z
                print('注入完成')
            else:
                # 填写坐标名称
                sht.range((1,col+1),(1,col+3)).options(transpose = False).value = name
                # 填写X坐标数据
                sht.range((2,col+1),(2+len(x),col+1)).options(transpose = True).value = x
                # 填写y坐标数据
                sht.range((2,col+2),(2+len(x),col+2)).options(transpose = True).value = y
                # 填写z坐标数据
                sht.range((2,col+3),(2+len(x),col+3)).options(transpose = True).value = z
                print('注入完成')
    finally:
        if wb:
            wb.save(args.input)
            wb.close()
            app.kill()
 """



























# 参数方程
""" theta1  = list(np.linspace(1.5*np.pi,2*np.pi,200))
x1 = 0 + R[4]*np.cos(theta1)
y1 = R[4] + R[4]*np.sin(theta1)
theta2  = np.linspace(np.pi/2,np.pi,200)
x2 = 360 + R[4]*np.cos(theta2)
y2 = (H[4]-R[4]) + R[4]*np.sin(theta2) """
# 第二种画圆的方法
""" plt.axis('equal')
plt.plot(x1,y1)
plt.plot(x2,y2)
plt.show()  """

# 第一种画圆的方法
""" # 画圆
fig = plt.figure(figsize=(1000,1000))
# plt.axis('equal')

for i in range(1):
    # 第一个圆
    plt.gca().set_xlim(-400,360)
    plt.gca().set_ylim(0,100)
    circle = plt.Circle((0, R[4]), R[4], color='y', fill=False)
    plt.gca().add_artist(circle)
    # 第二个圆
    circle = plt.Circle((360, H[4]-R[4]), R[4], color='r', fill=False)
    plt.gcf().gca().add_artist(circle)
    plt.show()
    plt.pause(0.05) """