# -*- encoding: utf-8 -*-
"""
test1.py
Created on 2018/10/29 15:24
@author: 马家树(majstx@163.com)
"""

import os

temp_files = 'C:/Users/meridian/Desktop/test/tmp'
source_file = 'C:/Users/meridian/Desktop/test/in.txt'
temp_path_list = []
if not os.path.exists(temp_files):
    os.makedirs(temp_files)
for i in range(0, 1000):

    temp_path_list.append(open(temp_files + str(i) + '.txt', mode='w'))
    print(temp_path_list[i])

with open(source_file) as f:
    for line in f:
        temp_path_list[hash(str(line)) % 1000].write(line)
        # print(hash(line) % 1000)
        # print(line)
for i in range(1000):
    print(temp_path_list[i])
    temp_path_list[i].close()
