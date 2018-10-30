# -*- encoding: utf-8 -*-
"""
cleanexcel.py
Created on 2018/8/21 13:36
Copyright (c) 2018/8/21, 
@author: 马家树(majstx@163.com)
"""
import openpyxl
from utils import myutil
import xlsxwriter
from utils.myutil import *

"""
按规则清除不可见字符
"""


def main(inpath, outpath):

    # 读
    wbr = openpyxl.load_workbook(inpath,  read_only=True)

    # 写入
    workbook = xlsxwriter.Workbook(filename=outpath)
    workbook.use_zip64()
    sheet = workbook.add_worksheet()

    rows = wbr.active.rows
    for row_num, row in enumerate(rows):
        printlog('当前处理的是第{0}行'.format(row_num))
        for col_num, cell in enumerate(row):
            val = cell.value
            val = myutil.str_formatxm(val)
            sheet.write(row_num, col_num, val)
    printlog('开始保存')
    workbook.close()
    printlog('保存成功')


if __name__ == '__main__':
    main()

