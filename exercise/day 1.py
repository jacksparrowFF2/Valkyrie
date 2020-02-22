#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   1.py
@Time    :   2020/02/15 21:51:28
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2020, EXphysiclab
@Desc    :   None
'''

# here put the import lib
from datetime import datetime,timedelta
# 获取今日日期
today = datetime.now()
# 输出今日日期
print(str(today))
# 获取年月日
year = today.year
month = today.month
day = today.day
# 输出年月日
print(year,month,day)
# 获取用户输入日期
birthday = input("please input your birthday:(dd/mm/yyyy)\n")
# 输出用户输入日期
birthday_date = datetime.strptime(birthday,'%d/%m/%Y')
print("your birthday is : " + str(birthday_date))

one_day = timedelta(days=1)
birthday_eve = birthday_date - one_day
print("your birthday eve is : " + str(birthday_eve))
