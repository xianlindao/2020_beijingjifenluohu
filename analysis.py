# Press the green button in the gutter to run the script.
import numpy as np
import pandas as pd
import seaborn
import prettytable as pt
from colorama import init, Fore, Back, Style
init(autoreset=False)

class Colored(object):
    #  前景色:红色  背景色:默认
    def red(self, s):
        return Fore.LIGHTRED_EX + s + Fore.RESET
    #  前景色:绿色  背景色:默认
    def green(self, s):
        return Fore.LIGHTGREEN_EX + s + Fore.RESET
    def yellow(self, s):
        return Fore.LIGHTYELLOW_EX + s + Fore.RESET
    def white(self,s):
        return Fore.LIGHTWHITE_EX + s + Fore.RESET
    def blue(self,s):
        return Fore.LIGHTBLUE_EX + s + Fore.RESET

from wcwidth import wcswidth

def get_aligned_string(string,width):
    string = "{:{width}}".format(string,width=width)
    bts = bytes(string,'utf-8')
    string = str(bts[0:width],encoding='utf-8',errors='backslashreplace')
    new_width = len(string) + int((width - len(string))/2)
    if new_width!=0:
        string = '{:{width}}'.format(str(string),width=new_width)
    return string

def read_excel(aPath):
    ## 读取 excel 文件
    return pd.read_excel(aPath)


def wc_rjust(text, length, padding=' '):
    from wcwidth import wcswidth
    text_len = wcswidth(text)
    if text_len + 4 >=12:
        return text

    return text + '\t'

if __name__ == '__main__':
    text = '北京'
    text_len = wcswidth(text)
    print('-' * (80 - text_len) + text)

    text = '哈尔滨'
    text_len = wcswidth(text)
    print('-' * (80 - text_len) + text)

    text = '乌鲁木齐'
    text_len = wcswidth(text)
    print('-' * (80 - text_len) + text)

    data = read_excel("./2020北京积分落户名单.xls")
    print(data.columns)
    print(data.values)
    print(data.query("创新创业 > 0")['单位名称'])

    data = data.head(20)

    name_max_len = 0
    data_name = data['姓名　　']
    for name in data_name:
        name_len = wcswidth(name)
        name_max_len = max(name_len, name_max_len)

    aa = '{:\u3000<%d}'%(name_len)
    f = lambda x: aa.format(x)
    data['姓名　　'] = data['姓名　　'].map(f)

    max_len = 0
    name = data['单位名称']
    for names in name:
        name_len = wcswidth(names)
        max_len = max(name_len, max_len)

    aa = '{:\u3000<%d}'%(max_len)
    f = lambda x: aa.format(x)
    data['单位名称'] = data['单位名称'].map(f)

    tb = pt.PrettyTable()
    tb.field_names = data.columns
    tb.align = "l"
    for row in data.values:
        row[1] = Fore.RED + row[1] + Fore.RESET
        tb.add_row(row)

    print(tb)

    # tmp_len = wcswidth('乌鲁木齐')
    #
    # x = pt.PrettyTable()
    # x.field_names = ["城市名称", "区号".ljust(6, ' '), "人口".ljust(6, ' '), "年降雨量".ljust(6, ' ')]
    # x.add_row(['{:\u3000<7}'.format('北京'), 1295, 1158259, 600.5])
    # x.add_row(['{:\u3000<7}'.format('哈尔滨'), 5905, 1857594, 1146.4])
    # x.add_row(['{:\u3000<7}'.format('上海'), 112, 120900, 1714.7])
    # x.add_row(['{:\u3000<7}'.format("乌鲁木齐"), 1357, 205556, 619.5])
    # print(x)

    print(data)