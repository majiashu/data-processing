# -*- encoding: utf-8 -*-
"""
xls2jsontest.py
Created on 2018/7/12 17:14
Copyright (c) 2018/7/12, 
@author: 马家树(majstx@163.com)
"""

import json
import xlrd

def readExcel():
    # 打开excel表单
    filename = u'信息 (2).xls'
    excel = xlrd.open_workbook(filename)

    # 得到第一张表单
    sheet1 = excel.sheets()[0]
    #找到有几列几列
    nrows = sheet1.nrows #行数
    ncols = sheet1.ncols #列数

    totalArray=[]
    title=[]
    # 标题
    for i in range(0,ncols):
        title.append(sheet1.cell(0,i).value)

    #数据
    for rowindex in range(1,nrows):
        dic={}
        for colindex in range(0,ncols):
            s=sheet1.cell(rowindex,colindex).value
            dic[title[colindex]]=s
        totalArray.append(dic)

    return json.dumps(totalArray,ensure_ascii=False)

print (readExcel())