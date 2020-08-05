# coding=UTF-8
import os


def formatDate(str_date):
    """
    用于格式化字符串
    :param str_date: 未格式化字符串
    :return: xxxx.xx.xx格式日期
    """
    split_parts = []
    # 2020.7.13
    if str_date.__contains__('.'):
        split_parts = str_date.split('.')
    # 2020-7-13
    elif str_date.__contains__('-'):
        split_parts = str_date.split('-')
    # 2020年7月13日
    elif str_date.__contains__('年'):
        tmp_year = str_date.split('年')[0]
        tmp_month = str_date.split('年')[1].split('月')[0]
        tmp_day = str_date.split('年')[1].split('月')[1].split('日')[0]
        split_parts.append(tmp_year)
        split_parts.append(tmp_month)
        split_parts.append(tmp_day)

    year = split_parts[0]
    month = split_parts[1].zfill(2)
    day = split_parts[2].zfill(2)
    return year + '.' + month + '.' + day


if __name__ == "__main__":
    start_time = format(input("Input the start time:\n"))
    end_time = format(input("Input the end time:\n"))
    os.system("python updateAll.py " + start_time + " " + end_time)
