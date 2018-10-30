# -*- encoding: utf-8 -*-
"""
fileclean.py
Created on 2018/8/9 18:06
Copyright (c) 2018/8/9
@author: 马家树(majstx@163.com)
"""

fr = open("C:/Users/meridian/Desktop/test/in.txt")
fw = open("C:/Users/meridian/Desktop/test/out.txt", "w")
lines = fr.readlines()
for line in lines:
    if '>>>>' in line:
        pass
    else:
        fw.write(line)





