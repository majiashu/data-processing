# -*- encoding: utf-8 -*-
"""
xls2json2.0.py
Created on 2018/7/17 13:27
Copyright (c) 2018/7/17, 
@author: 马家树(majstx@163.com)
"""
from collections import OrderedDict
import xlrd
from datetime import datetime
from xlrd import xldate_as_tuple


# wb = xlrd.open_workbook("协和医院核医学科分化型甲状腺癌数据录入表格结构---2018-07-10.xls")
# sh = wb.sheet_by_index(0)


def getinfo():
    """
    该方法处理基本信息、问卷部分
    :return: dict_jbxx
    """
    dict_hb = {}
    dict_temp1 = {}  # OrderedDict()
    dict_temp2 = {}  # OrderedDict()
    dict_temp3 = {}
    dict_temp4 = {}

    # 获取基本信息字典
    for colnum in range(0, 12):
        # print(colnum)
        dic_key = sh.cell(1, colnum).value
        # print(dic_key)
        dic_val = sh.cell(3, colnum).value
        # 判断cell格式并格式化
        c_type = sh.cell(3, colnum).ctype
        if c_type == 2:
            dic_val = int(dic_val)
        elif c_type == 3:
            date = datetime(*xldate_as_tuple(dic_val, 0))
            dic_val = date.strftime('%Y/%d/%m')

        dict_temp1[dic_key] = dic_val
    # print(dict_temp1)

    # 获取问卷信息字典
    for colnum in range(12, 24):
        dic_key = sh.cell(1, colnum).value
        dic_val = sh.cell(3, colnum).value
        # 判断cell格式并格式化
        c_type = sh.cell(3, colnum).ctype
        if c_type == 2:
            dic_val = int(dic_val)
        elif c_type == 3:
            date = datetime(*xldate_as_tuple(dic_val, 0))
            dic_val = date.strftime('%Y/%d/%m')

        dict_temp2[dic_key] = dic_val

    # 获取手术前甲功
    for colnum in range(24, 30):
        dic_key = sh.cell(1, colnum).value
        dic_val = sh.cell(3, colnum).value
        # 判断cell格式并格式化
        c_type = sh.cell(3, colnum).ctype
        if c_type == 3:
            date = datetime(*xldate_as_tuple(dic_val, 0))
            dic_val = date.strftime('%Y/%d/%m')
        dict_temp3[dic_key] = dic_val

    # 获取STAGE1(7TH)
    stage1_val = sh.cell(3, 76).value

    # 获取STAGE2(8TH)
    stage2_val = sh.cell(3, 77).value

    # 获取术后距离碘-131治疗时间（单位：天）
    shjldts = sh.cell(3, 78).value
    shjldts = int(shjldts)

    # 获取术后TSH抑制治疗检验结果
    for colnum in range(79, 97):
        dic_key = sh.cell(1, colnum).value
        dic_val = sh.cell(3, colnum).value
        dict_temp4[dic_key] = dic_val

    # 获取碘131治疗前评估
    dict_temp5 = {}
    for colnum in range(97, 122):
        dic_key = sh.cell(1, colnum).value
        dic_val = sh.cell(3, colnum).value
        # 判断cell格式并格式化
        c_type = sh.cell(3, colnum).ctype
        if c_type == 3:
            date = datetime(*xldate_as_tuple(dic_val, 0))
            dic_val = date.strftime('%Y/%d/%m')
        dict_temp5[dic_key] = dic_val

    # 获取碘-131治疗
    dict_temp6 = {}
    for colnum in range(122, 127):
        dic_key = sh.cell(1, colnum).value
        dic_val = sh.cell(3, colnum).value
        # 判断cell格式并格式化
        c_type = sh.cell(3, colnum).ctype
        if c_type == 3:
            date = datetime(*xldate_as_tuple(dic_val, 0))
            dic_val = date.strftime('%Y/%d/%m')
        dict_temp6[dic_key] = dic_val

    # 碘－131治疗后随诊
    val_list = []
    for r in range(3, 9):
        dic_temp = {}
        for colnum in range(127, 146):
            dic_key = sh.cell(1, colnum).value
            dic_val = sh.cell(r, colnum).value
            # 判断cell格式并格式化
            c_type = sh.cell(r, colnum).ctype
            if c_type == 3:
                date = datetime(*xldate_as_tuple(dic_val, 0))
                dic_val = date.strftime('%Y/%d/%m')

            dic_temp[dic_key] = dic_val
        # print(dic_temp)
        val_list.append(dic_temp)
        # print(val_list)

    dict_hb["患者基本信息（确诊时状态）"] = dict_temp1
    dict_hb["问卷信息"] = dict_temp2
    dict_hb["手术前甲功"] = dict_temp3
    dict_hb["STAGE1(7TH)"] = stage1_val
    dict_hb["获取STAGE2(8TH)"] = stage2_val
    dict_hb["术后距离碘-131治疗时间（单位：天）"] = shjldts
    dict_hb["术后TSH抑制治疗检验结果"] = dict_temp4
    dict_hb["碘131治疗前评估"] = dict_temp5
    dict_hb["碘-131治疗"] = dict_temp6
    dict_hb["碘－131治疗后随诊"] = val_list
    # print(dict_hb)
    return dict_hb


# def d131dinfo():    # cell(127, 146)
#     dict_hb2 = {}
#     val_list = []
#     for r in range(3, 9):
#         dic_temp = {}
#         for colnum in range(127, 146):
#             dic_key = sh.cell(1, colnum).value
#             dic_val = sh.cell(r, colnum).value
#             # 判断cell格式并格式化
#             c_type = sh.cell(r, colnum).ctype
#             if c_type == 3:
#                 date = datetime(*xldate_as_tuple(dic_val, 0))
#                 dic_val = date.strftime('%Y/%d/%m')
#
#             dic_temp[dic_key] = dic_val
#         # print(dic_temp)
#         val_list.append(dic_temp)
#     # print(val_list)
#     dict_hb2["碘－131治疗后随诊"] = val_list
#     print(dict_hb2)


if __name__ == '__main__':
    wb = xlrd.open_workbook("协和医院核医学科分化型甲状腺癌数据录入表格结构---2018-07-10.xls", formatting_info=True)
    sh = wb.sheet_by_index(0)
    # 获取合并单元格的边界
    mergcell_list = sh.merged_cells
    bj_list = []
    for num_tup in range(len(mergcell_list)):
        val_tup = mergcell_list[num_tup]
        if (val_tup[2] == 0 and val_tup[3] == 1):
            bj_list.append(val_tup)
    for bj in range(1, len(bj_list)):
        bj_tup = bj_list[bj]
        bj_up = bj_tup[0]
        bi_down = bj_tup[1]
        print(bj_tup, bj_up, bi_down)
    dic = getinfo()
    print(dic)


