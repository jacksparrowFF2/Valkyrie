#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   AutoPixel.py
@Time    :   2021/08/05 21:05:13
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2021, EXphysiclab
@Desc    :   None
'''

# here put the import lib
from pyautocad import Autocad,APoint
import time

from pyautocad.types import aDouble

# 定义像素相关信息
pitch = 2.5
double = pitch*2
# AA屏的长和宽
AA_length = 50
AA_height = 50
# 确定循环次数
loop_x = round(AA_length/(2*pitch))
print(loop_x)
loop_y = round(AA_height/(2*pitch))
print(loop_y)

# 第一排
# 生成 RG 横坐标序列
RG_x = [2*pitch * x for x in range(0,loop_x+1)]
print(RG_x)
# 生成 BG 横坐标序列
BG_x = [pitch * x for x in range(0,loop_x+1)]
# 生成第一排的纵坐标
y1 = [-2.5 * x for x in range(1,loop_y+1)]
print(BG_x)

# 第二排
# 生成 BG-2 横坐标序列
# 减少性能占用，每隔 0.5 s进行一次循环

# # 在坐标(x,y)插入块
# insertionPnt = APoint(0,-2.5)
# RetVal = acad.model.InsertBlock(insertionPnt,"D:\Program Files\github graduate\Valkyrie\Vision\CAD\RG.dwg",1,1,1,0)
# # 延迟 0.5 s
# time.sleep(0.5)
# # 在坐标(x,y)插入块
# insertionPnt = APoint(2.5,-2.5)
# RetVal = acad.model.InsertBlock(insertionPnt,"D:\Program Files\github graduate\Valkyrie\Vision\CAD\BG.dwg",1,1,1,0)
# # 延迟 0.5 s
# time.sleep(0.5)
# # 在坐标(x,y)插入块
# insertionPnt = APoint(2.5,-5)
# RetVal = acad.model.InsertBlock(insertionPnt,"D:\Program Files\github graduate\Valkyrie\Vision\CAD\RG-2.dwg",1,1,1,0)
# # 延迟 0.5 s
# time.sleep(0.5)
# # 在坐标(x,y)插入块
# insertionPnt = APoint(0,-5)
# RetVal = acad.model.InsertBlock(insertionPnt,"D:\Program Files\github graduate\Valkyrie\Vision\CAD\BG-2.dwg",1,1,1,0)
# # 延迟 0.5 s
# time.sleep(0.5)

# 连接CAD
acad = Autocad(create_if_not_exists = True)
acad.prompt("Hello! AutoCAD from pyautocad.")
print(acad.doc.Name)
acadmod = acad.ActiveDocument.ModelSpace  # 图形空间 方法1





# 在（x,y）位置处创建矩形 
# 定义矩形的坐标点
p1 = APoint(0.0)
p2 = APoint(2*pitch,0)
p3 = APoint(2*pitch,-2*pitch)
p4 = APoint(0,-2*pitch)
# 连接成线,创建矩形
Pixel_2pitch = [p1,p2,p3,p4,p1]
# 将各点坐标顺序变换为1行多列的1维数组
Pixel_2pitch = [j for i in Pixel_2pitch for j in i]
print(Pixel_2pitch)
# 将数据类型转换为双精度浮点数
Pixel_2pitch = aDouble(Pixel_2pitch)
# 绘制图形
Pixel_Cell = acad.model.AddPolyLine(Pixel_2pitch)


# 阵列对象
# # 矩型阵列
# numberOfRows = 3
# numberOfColumns = 3
# numberOfLevels = 1
# distanceBwtnRows = 5
# distanceBwtnColumns = 5
# distanceBwtnLevels = 1
# try:
#     retObj = ciobj1.ArrayRectangular (numberOfRows, numberOfColumns, numberOfLevels,
#                                 distanceBwtnRows, distanceBwtnColumns, distanceBwtnLevels)
# except:
#     pass

