#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   test_data_extract.py
@Time    :   2021/07/31 20:34:26
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2021, EXphysiclab
@Desc    :   None
'''

# here put the import lib
import os
import argparse
import win32clipboard
import win32con
import numpy as np
import xlwings as xw
import math

parser = argparse.ArgumentParser(description="该脚本用于对测试数据型号1数据进行整合并处理，\
    搭配Quicker进行使用效果更佳")