
#!/usr/bin/python3
# -*- coding: UTF8 -*-

import win32clipboard as wc
import win32con
import chardet
import ast

#获取剪贴板内容


def getCopyText():
    wc.OpenClipboard()
    copy_text = wc.GetClipboardData(win32con.CF_UNICODETEXT)
    wc.CloseClipboard()
    return copy_text


#将字符串类型转别为字典类型
excode = ast.literal_eval(getCopyText())

# print(excode)
print(type(excode))
print(excode["日期"])
# for key, value in excode.items():
#     print(key, value)

date = int(excode["日期"])
power = int(excode["初始输入功率(w)"])-int(excode["初始反馈功率(w)"])
Ar = int(excode["Ar(sccm)"])
H2 = int(excode["H2(sccm)"])
CH4 = int(excode["CH4(sccm)"])
pressure = int(excode["CH4(sccm)"])
temp = int(excode["温度(℃)"])
sub1 = excode["衬底1"]
sub2 = excode["衬底2"]
metaltype = excode["金属网"]
note = excode["实验目的"]

print(date, power, Ar, H2, CH4, pressure, temp, sub1+sub2, metaltype, note)
