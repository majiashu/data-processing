# -*- encoding: utf-8 -*-
"""
myutil.py
Created on 2018/1/3 13:58
Copyright (c) 2018/1/3, 
@author: 小马同学(majstx@163.com)
"""
import re
import datetime
import decimal
import time

_clean_strs = (
    '\xe3\x80\x80',  #
    '\x1c',  #
    '\x1d',  #
    '\x3f'
)


def number2char(number):
    high_chr = ''
    if number == 0:
        raise Exception('"0" is a bad imput')
    if number > 26 and number % 26 == 0:
        high_chr = number2char(number // 26 - 1)
    elif number > 26:
        high_chr = number2char(number // 26)
    if number % 26 == 0:
        low_chr = chr(64 + 26)
    else:
        low_chr = chr(64 + number % 26)
    return high_chr + low_chr


# 比较方法将单元格内容改为 正常 偏高 偏低
def compare(jg_value, range_value):
    range_value = range_value.replace('--', '-').replace('~', '-').replace('－', '-').replace('～', '-').replace('ー',
                                                                                                               '-').replace(
        '%', '')

    if range_value == '-':
        compare_value = jg_value
    elif '-' in range_value:
        range_low = range_value.split('-')[0]
        range_low = float(range_low)
        range_hight = range_value.split('-')[1]
        range_hight = float(range_hight)
        if float(jg_value) < range_low:
            compare_value = '偏低'
        elif float(jg_value) > range_hight:
            compare_value = '偏高'
        else:
            compare_value = '正常'
    elif '>' in range_value:
        ljz_value = range_value[1:]
        if float(jg_value) > float(ljz_value):
            compare_value = '正常'
        else:
            compare_value = '偏低'
    elif '<' in range_value:
        ljz_value = range_value[1:]
        if float(jg_value) < float(ljz_value):
            compare_value = '正常'
        else:
            compare_value = '偏高'
    else:
        compare_value = jg_value
    return compare_value


# 判断字符串是否为数字
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False


def tostr(value):
    if value is not None:
        value_type = type(value)
        tuple_2str = (datetime.datetime, datetime.date, float, int, decimal.Decimal, bool, list, tuple, set, dict)
        if value_type == str:
            value.strip()
        # elif value_type == unicode:
        #     if value == u'\x00':
        #         value = ''
        #     else:
        #         value = value.encode('UTF-8').strip()
        elif value_type in tuple_2str:
            value = str(value).strip()
        else:
            print('{} >>>>> {}'.format(value, value_type))
            value = ''
    else:
        value = ''
    return value


# def tounicode(value):
#     if value is not None:
#         value_type = type(value)
#         if value_type != unicode:
#             value = tostr(value)
#             value = value.decode('UTF-8')
#     else:
#         value = u''
#     return value.strip()

def str_formatxm(value):
    """
    1) 将首字符为 '=' 的字符串进行处理，防止在Excel中出现公式错误
    2) 清洗显示为乱码的字符
    3) 不管有几个空格保留一个  ' '.join(s.split())
    4) 换行符变为中文句号，多余一个的中文句号保留一个
    :param value:
    :return: value
    """
    value = tostr(value)
    if value.find('=') == 0:
        value = '＝' + value[1:]
    value = value.replace('\t', '').replace('\r', '').replace('\n', '。').replace('_x000D_', '')
    value = re.sub(' +', ' ', value)
    # value = tounicode(value)
    value = re.sub(u'。+', u'。', value)
    # value = value.encode('utf-8')
    value = tostr(value)
    for c in _clean_strs:
        if c in value:
            value = value.replace(c, '')
    # value = confir(value)
    return value


# 去除不可见字符
def confir(strc):
    for i in range(0, 32):
        strc = strc.replace(chr(i), '')
    return strc


# 打印带有当前时间
def printlog(agrs):
    print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()), agrs)
