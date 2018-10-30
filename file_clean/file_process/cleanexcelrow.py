# -*- encoding: utf-8 -*-
"""
cleanexcel.py
Created on 2018/8/21 13:36
Copyright (c) 2018/8/21, 
@author: 马家树(majstx@163.com)
"""
import openpyxl
import xlsxwriter
from utils.myutil import *

"""
按规则清除不可见字符，read_only=Ture模式 一行一行处理 节省内存资源
"""


def main():
    # 读
    wbr = openpyxl.load_workbook('C:/Users/meridian/Desktop/襄阳航空工业医院/襄阳航空工业/exportdata_classifyt.xlsx', read_only=True)
    rows = wbr.active.rows

    # 写入
    workbook = xlsxwriter.Workbook(filename='C:/Users/meridian/Desktop/襄阳航空工业医院（新）/exportdata_classify.xlsx')
    workbook.use_zip64()
    sheet = workbook.add_worksheet()

    for row_num, row in enumerate(rows):
        if row_num % 500 == 0:
            printlog('当前处理的是第{0}行'.format(row_num))
        for col_num, cell in enumerate(row):
            val = cell.value
            if col_num > 15:  # column start num
                val = str_formatxm(val)
            sheet.write(row_num, col_num, val)
    printlog('开始保存')
    workbook.close()
    printlog('保存成功')


if __name__ == '__main__':
    main()
