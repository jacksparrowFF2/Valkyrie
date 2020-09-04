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
        # 切片次数
        t = 10
        # 生成 z 轴坐标序列
        Z = list(range(0, t, 1))
        print(Z)
        # 生成高度序列
        H = list(np.linspace(0.000000000000001,40,t))
        print(H)
        # 生成半径序列
        R = [x/4 + (360*90)/x for x in H]
        # 限定 X 轴坐标
        x1= np.linspace(0,180,360)
        temp = np.linspace(-180,0,360)
        x2 = np.linspace(180,360,360)
        x = list(x1) + list(x2)
        # print(x)
        m = 0
        for i in range(len(H)):
            z = [Z[m]]*len(x)
            # print(z)
            m = m + 1
            # 一般方程
            y1 = R[i]-np.sqrt(np.square(R[i])-np.square(x1))
            y2 = (H[i]-R[i])+np.sqrt(np.square(R[i])-np.square(temp))
            y = list(y1) + list(y2)
            # print(y)
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

        """     # 画图
            plt.axis('equal')
            plt.plot(x1,y1)
            plt.plot(x2,y2)
            plt.show()
            plt.pause(0.05) """
    finally:
        if wb:
            wb.save(args.input)
            wb.close()
            app.kill()



































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