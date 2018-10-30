# -*- encoding: utf-8 -*-
"""
gettop1000.py
Created on 2018/9/13 16:28
@author: 马家树(majstx@163.com)
"""

import openpyxl
import time
import pandas as pd

start = time.time()

# wb = openpyxl.load_workbook('C:/Users/meridian/Desktop/九江市第一人民医院/九江市第一人民医院_原始.xlsx')
data = pd.read_excel('C:/Users/meridian/Desktop/九江市第一人民医院/九江市第一人民医院_原始.xlsx', sheet_name=0)

end = time.time()
use_time = end - start
print("读取该文件耗时{0}".format())
