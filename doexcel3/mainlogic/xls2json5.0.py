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
import json
import codecs

"""
术后TSH抑制治疗检验结果   将手术前第一次与最后一次的数据分开
"""


def getinfo(up, down):
    """
    该方法处理基本信息、问卷部分
    :return: dict_jbxx
    """
    dict_hb = OrderedDict()

    # 获取基本信息字典
    dict_temp1 = OrderedDict()
    for colnum in range(0, 12):
        # print(colnum)
        dic_key = sh.cell(1, colnum).value
        # print(dic_key)
        dic_val = sh.cell(up, colnum).value
        # 判断cell格式并格式化
        c_type = sh.cell(up, colnum).ctype
        if c_type == 2:
            dic_val = int(dic_val)
        elif c_type == 3:
            date = datetime(*xldate_as_tuple(dic_val, 0))
            dic_val = date.strftime('%Y/%d/%m')

        dict_temp1[dic_key] = dic_val
    # print(dict_temp1)

    # 获取问卷信息字典
    dict_temp2 = OrderedDict()
    for colnum in range(12, 24):
        dic_key = sh.cell(1, colnum).value
        dic_val = sh.cell(up, colnum).value
        # 判断cell格式并格式化
        c_type = sh.cell(up, colnum).ctype
        if c_type == 2:
            dic_val = int(dic_val)
        elif c_type == 3:
            date = datetime(*xldate_as_tuple(dic_val, 0))
            dic_val = date.strftime('%Y/%d/%m')

        dict_temp2[dic_key] = dic_val

    # 获取手术前甲功
    dict_temp3 = OrderedDict()
    for colnum in range(24, 30):
        dic_key = sh.cell(1, colnum).value
        dic_val = sh.cell(up, colnum).value
        # 判断cell格式并格式化
        c_type = sh.cell(up, colnum).ctype
        if c_type == 3:
            date = datetime(*xldate_as_tuple(dic_val, 0))
            dic_val = date.strftime('%Y/%d/%m')
        dict_temp3[dic_key] = dic_val
    # 获取手术及病理情况
    val_list0 = []
    for r in range(up, down):
        dic_temp = OrderedDict()
        flag_val = sh.cell(r, 30).value
        print(type(flag_val))
        if (flag_val is None) or (flag_val == ""):
            continue
        for colnum in range(30, 35):
            dic_key = sh.cell(1, colnum).value
            dic_val = sh.cell(r, colnum).value
            # 判断cell格式并格式化
            c_type = sh.cell(r, colnum).ctype
            if c_type == 3:
                date = datetime(*xldate_as_tuple(dic_val, 0))
                dic_val = date.strftime('%Y/%d/%m')
            dic_temp[dic_key] = dic_val
        dic_tnm7 = OrderedDict()
        for colnum in range(35, 38):
            dic_key = sh.cell(2, colnum).value
            dic_val = sh.cell(r, colnum).value
            dic_tnm7[dic_key] = dic_val
        dic_temp["TNM分期（AJCC第七版）"] = dic_tnm7
        dic_tnm8 = OrderedDict()
        for colnum in range(38, 41):
            dic_key = sh.cell(2, colnum).value
            dic_val = sh.cell(r, colnum).value
            dic_tnm8[dic_key] = dic_val
        dic_temp["TNM分期（AJCC第八版）"] = dic_tnm8
        for colnum in range(41, 46):
            dic_key = sh.cell(1, colnum).value
            dic_val = sh.cell(r, colnum).value
            dic_temp[dic_key] = dic_val
        dic_zclbjqs = OrderedDict()
        for colnum in range(46, 52):
            dic_key = sh.cell(2, colnum).value
            dic_val = sh.cell(r, colnum).value
            dic_zclbjqs[dic_key] = dic_val
        dic_temp["左侧淋巴结清扫"] = dic_zclbjqs
        dic_yclbjqs = OrderedDict()
        for colnum in range(53, 58):
            dic_key = sh.cell(2, colnum).value
            dic_val = sh.cell(r, colnum).value
            dic_yclbjqs[dic_key] = dic_val
        dic_temp["右侧淋巴结清扫"] = dic_yclbjqs
        dic_key = sh.cell(1, 58).value
        dic_val = sh.cell(r, 58).value
        dic_temp[dic_key] = dic_val
        dic_zczysl = OrderedDict()
        for colnum in range(59, 65):
            dic_key = sh.cell(2, colnum).value
            dic_val = sh.cell(r, colnum).value
            dic_zczysl[dic_key] = dic_val
        dic_temp["左侧转移数量"] = dic_zczysl
        dic_yczysl = OrderedDict()
        for colnum in range(65, 71):
            dic_key = sh.cell(2, colnum).value
            dic_val = sh.cell(r, colnum).value
            dic_yczysl[dic_key] = dic_val
        dic_temp["右侧转移数量"] = dic_yczysl
        for colnum in range(71, 76):
            dic_key = sh.cell(1, colnum).value
            dic_val = sh.cell(r, colnum).value
            dic_temp[dic_key] = dic_val
        val_list0.append(dic_temp)

    # 获取STAGE1(7TH)
    stage1_val = sh.cell(up, 76).value

    # 获取STAGE2(8TH)
    stage2_val = sh.cell(up, 77).value

    # 获取术后距离碘-131治疗时间（单位：天）
    shjldts = sh.cell(up, 78).value
    shjldts = int(shjldts)

    # 获取术后TSH抑制治疗检验结果
    # dict_temp4 = OrderedDict()
    # for colnum in range(79, 97):
    #     dic_key = sh.cell(1, colnum).value
    #     dic_val = sh.cell(up, colnum).value
    #     dict_temp4[dic_key] = dic_val
    dic_shtsh = OrderedDict()  # 小总

    dic_shtysta = OrderedDict()
    for colnum in range(79, 88):
        dic_key = sh.cell(2, colnum).value
        dic_val = sh.cell(up, colnum).value
        dic_shtysta[dic_key] = dic_val
    dic_shtsh["术后停药前第一次检验"] = dic_shtysta

    dic_shtyend = OrderedDict()
    for colnum in range(88, 97):
        dic_key = sh.cell(2, colnum).value
        dic_val = sh.cell(up, colnum).value
        dic_shtyend[dic_key] = dic_val
    dic_shtsh["术后停药前最后一次检验"] = dic_shtyend

    # 获取碘131治疗前评估
    dict_temp5 = OrderedDict()
    for colnum in range(97, 122):
        dic_key = sh.cell(1, colnum).value
        dic_val = sh.cell(up, colnum).value
        # 判断cell格式并格式化
        c_type = sh.cell(up, colnum).ctype
        if c_type == 3:
            date = datetime(*xldate_as_tuple(dic_val, 0))
            dic_val = date.strftime('%Y/%d/%m')
        dict_temp5[dic_key] = dic_val

    # 获取碘-131治疗
    dict_temp6 = OrderedDict()
    for colnum in range(122, 127):
        dic_key = sh.cell(1, colnum).value
        dic_val = sh.cell(up, colnum).value
        # 判断cell格式并格式化
        c_type = sh.cell(up, colnum).ctype
        if c_type == 3:
            date = datetime(*xldate_as_tuple(dic_val, 0))
            dic_val = date.strftime('%Y/%d/%m')
        dict_temp6[dic_key] = dic_val

    # 碘－131治疗后随诊
    val_list = []
    for r in range(up, down):
        dic_temp = OrderedDict()
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
    dict_hb["手术及病理情况"] = val_list0
    dict_hb["STAGE1(7TH)"] = stage1_val
    dict_hb["STAGE2(8TH)"] = stage2_val
    dict_hb["术后距离碘-131治疗时间（单位：天）"] = shjldts
    dict_hb["术后TSH抑制治疗检验结果"] = dic_shtsh
    dict_hb["碘131治疗前评估"] = dict_temp5
    dict_hb["碘-131治疗"] = dict_temp6
    dict_hb["碘－131治疗后随诊"] = val_list
    # print(dict_hb)
    return dict_hb


if __name__ == '__main__':
    wb = xlrd.open_workbook("协和医院核医学科分化型甲状腺癌数据录入表格结构---2018-07-10 - 副本.xls", formatting_info=True)
    sh = wb.sheet_by_index(0)
    json_list = []
    # 获取合并单元格的边界
    mergcell_list = sh.merged_cells
    bj_list = []
    for num_tup in range(len(mergcell_list)):
        val_tup = mergcell_list[num_tup]
        if (val_tup[2] == 0 and val_tup[3] == 1):
            bj_list.append(val_tup)
    for bj in range(1, len(bj_list)):
        bj_up, bi_down = bj_list[bj][0], bj_list[bj][1]
        print(bj_up, bi_down)
        dic = getinfo(bj_up, bi_down)
        json_list.append(dic)

    j = json.dumps(json_list, ensure_ascii=False)
    with codecs.open('test1.json', "w", "utf-8") as f:
        f.write(j)


