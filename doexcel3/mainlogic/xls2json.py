# -*- encoding: utf-8 -*-
"""
xls2json.py
Created on 2018/7/11 13:25
Copyright (c) 2018/7/11, 
@author: 马家树(majstx@163.com)
"""

import xlrd
from collections import OrderedDict
import json
import codecs
from datetime import datetime
from xlrd import xldate_as_tuple

wb = xlrd.open_workbook('test.xlsx')
convert_list = []
sh = wb.sheet_by_index(2)
title = sh.row_values(0)
for rownum in range(1, sh.nrows):
    rowvalue = sh.row_values(rownum)  # 获取excel一行数据返回一个列表
    single = OrderedDict()
    for colnum in range(0, len(rowvalue)):
        # print(title[colnum], rowvalue[colnum])
        colvalue = rowvalue[colnum]
        c_type = sh.cell(rownum, colnum).ctype
        print(c_type)
        if c_type == 2:
            colvalue = int(colvalue)
        elif c_type == 3:
            colvalue = sh.cell(rownum, colnum).value
            date = datetime(*xldate_as_tuple(colvalue, 0))
            colvalue = date.strftime('%Y/%d/%m')
        single[title[colnum]] = colvalue
    convert_list.append(single)

print(convert_list)

j = json.dumps(convert_list)
print(j)

with codecs.open('file.json', "w", "utf-8") as f:
    f.write(j)
