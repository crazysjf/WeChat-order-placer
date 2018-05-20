# -*- coding: utf-8 -*-
from datetime import date

def gen_order_text(orders, p):
    '生成单个档口的报单文本'
    text = ''
    try:
        t = date.today()  # 仅获取日期
        text = u'报单\n日期：%s月%s日\n档口：%s\n\n' % (t.month, t.day, p)
        text = text + "------------------------------\n"
        o = orders[p]
        last_code = 'dummy'
        for i,l in enumerate(o):
            code = l['code']
            s = ""
            # 分隔款式
            if code != last_code and i != 0:
                s = s + "------\n"

            s = s + "%-5s,\t%-10s,\t%-5s\n" % (code, l['spec'] , l['nr'])
            text = text + s
            last_code = code
        text = text + "------------------------------\n\n"
        text = text + u" - 为了大家的方便，请开实价、开实数，避免欠货。\n"
        text = text + u" - 网店经营，颜色、尺码缺货请不要拼凑。\n"
        text = text + u" - 请微信拍照发单，这边2小时内付款，货由工仔收。\n"
        text = text + u" - 合作愉快！"



    except:
        text = '报单生成错误'
    return text



def gen_all_orders_text(orders):
    s = u''
    for p in orders.keys():
        s = s + gen_order_text(orders, p)
        s = s + '\n'
    return s