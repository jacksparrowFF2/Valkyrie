#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   text.py
@Time    :   2019/09/16 21:43:29
@Author  :   SPH 
@Version :   1.0
@Contact :   s.ph@outlook.com
@License :   (C)Copyright 2018-2019, EXphysiclab
@Desc    :   None
'''

# here put the import lib

import argparse

parser = argparse.ArgumentParser(description='This program is aims to extract experiment data from txt file')
# parser.add_argument("-i","--input",metavar="", type = argparse.FileType("r"), required = True, 
#                     help = "txt file to process")
# parser.add_argument("-i","--input",metavar="", type = str, required = True, 
#                     help = "txt file to process")
parser.add_argument("-i","--input", metavar= "INFILE", type = argparse.FileType(mode = 'r'), 
                    help = 'the file need to extract data')

group = parser.add_mutually_exclusive_group()
group.add_argument('-q','--quiet', action='store_true',help = 'print quiet')
group.add_argument('-v','--verbose', action = 'store_true', help = 'print verbose')
args = parser.parse_args()

""" if __name__ == '__main__':
    a = args.input.read()
    if args.quiet:
      print(a)
      print(type(a))
    elif args.verbose:
      print("Your input filename is %s" %(args.input))
    else:
        print(a) """
c = []
if __name__ == '__main__':
    infile = args.input
    b = infile.read()
    if args.quiet:
      print(b)
    elif args.verbose:
      print("Your input filename is %s" %(args.input))
    else:
        print(b)