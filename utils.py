# -*- coding: utf-8 -*-
from datetime import date

def gen_order_text(orders, p):
    '生成单个档口的报单文本'
    text = ''
    try:
        t = date.today()  # 仅获取日期
        text = u'报单(网店史小姐)\n日期：%s月%s日\n档口：%s\n\n' % (t.month, t.day, p)
        text = text + "------------------------------\n"
        o = orders[p]
        for l in o:
            s = "%-5s,\t%-10s,\t%-5s\n" % (l['code'], l['spec'] , l['nr'])
            text = text + s
        text = text + "------------------------------\n"
    except:
        text = '报单生成错误'
    return text



def print_all_order_text(orders):
    for p in orders.keys():
        print gen_order_text(orders, p)
        print