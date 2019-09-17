#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   get_txt.py
@Time    :   2019/09/16 20:52:02
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import win32clipboard as w
import win32con
import win32api

# name = input()
# print(type(name))

d = {}

with open(r'F:\github_graduate\Valkyrie\idea\201909 04.txt', 'r',encoding='utf-8') as f:
    # line = f.readline()
    a = f.read()
# n = line.split(' ')
# print(line)
# print(len(n))
print(a)

# b = a.split(' ')




b = a.split('\n')
print(b)
print(type(b))

n = print(len(b))

c = []
for i in b:
    c.append(i.split('\t'))   
print(c)
# print(c[0][0])
# d = c[0][0].split(' ')
# print(d[1])


d = {}
for i in range(len(b)):
    d[c[i][0]] = c[i][1]
print(d)
print(type(d))
e = list(d.values())
print(e)
print(type(e))
f = str(e)
print(f)
print(type(f))
g = f.replace('\'', ' ')
print(g)
g = g.replace(',', '\n')
print(g)
g = g.replace('[', ' ')
print(g)
g = g.replace(']', ' ')
print(g)


# print ("Value : %s" %  d.values())

""" b = "sss"
# 写入剪贴板
print(type(b))

def writeclip(astring):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, astring)
    w.CloseClipboard()

writeclip(a) """