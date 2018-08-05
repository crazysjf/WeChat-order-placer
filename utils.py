# -*- coding: utf-8 -*-
from datetime import date
import re

def gen_order_text(orders, p):
    '''
    生成单个档口的报单文本
    
    :param orders: 报单字典，包括所有档口数据
    :param p: 所要生成报单文本的档口
    :return: 
    '''
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
        text = text + u" - 请微信拍照发单，2小时内付款，工仔收货。\n"
        text = text + u" - 该单仅当日有效。合作愉快！"



    except:
        text = '报单生成错误'
    return text

def print_exception_summary(r):
    '''
    打印到货异常汇总。
    :param r: 异常字典
    :return: 
    '''
    print("\n拿货异常：")
    print("-------------------")

    e_cnt = 0
    for p in r.keys():
        lines = r[p]
        for l in lines:
            e_cnt = e_cnt + 1
            print("%s, %s, %s, %s" % (p, l['code'], l['spec'], l['nr']))
        print("-------------------")
    print("档口数： % s，异常数： % s\n" % (len(r.keys()), e_cnt))


def gen_all_orders_text(orders):
    s = u''
    for p in orders.keys():
        s = s + gen_order_text(orders, p)
        s = s + '\n'
    return s


def gen_exception_text(e, p):
    '''
     生成单个档口的到货异常文本

     :param orders: 到货异常字典，包括所有档口数据
     :param p: 所要生成文本的档口
     :return: 
     '''
    text = ''
    try:
        t = date.today()  # 仅获取日期
        text = u'到货及欠货确认\n日期：%s月%s日\n档口：%s\n\n' % (t.month, t.day, p)
        text = text + "------------------------------\n"
        o = e[p]
        last_code = 'dummy'
        for i, l in enumerate(o):
            code = l['code']
            s = ""
            # 分隔款式
            if code != last_code and i != 0:
                s = s + "------\n"

            s = s + "%-5s,\t%-10s,\t到%s件，欠%s件\n" % (code, l['spec'], l['total'] - l['nr'], l['nr'])
            text = text + s
            last_code = code
        text = text + "------------------------------\n\n"
        text = text + " - 有异议请及时回复\n"
        text = text + " - 无异议无需回复\n"

    except:
        text = '报单生成错误'
    return text

def gen_all_exception_text(e):
    s = ''
    for p in e.keys():
        s = s + gen_exception_text(e, p)
        s = s + '\n'
    return s


def convert_possible_num_to_str(v):
    '如果v是数字类型，则转为字符串，否则原样返回'
    t = type(v).__name__
    if t == 'float' or t == 'long' or  t == "int":
        return str(int(v))
    else:
        return v


def get_store(all_friends, code):
    '''
    itchat提供的search_friends并没有模糊匹配的功能，需要自己实现。
    根据code商家编码来查找好友并返回。
    未找到返回None。
    TODO：需要处理多个匹配的情况
    只要all_friends里面有好友的remarkName前面部分和这个匹配(忽略大小写)，即视为找到
    all_friends: itchat.get_friends()获取的所有好友的列表
    code：商家编码，如11152-茉莉
    '''
    for f in all_friends:
        try:
            pat = r'^%s' % code
        except Exception as e:
            print("e.message: %s" % code)
        remarkName = f['RemarkName']
        nickName = f['NickName']
        #print f
        #print remarkName
        m1 = re.match(pat, remarkName, re.IGNORECASE)
        m2 = re.match(pat, nickName, re.IGNORECASE)

        if m1 != None or m2 != None:
            # 备注和昵称两个左右匹配了一个就算找到
            return f

    return None


def print_friend(f):
    if f == None:
        print('无此好友')
        return
    for k in f.keys():
        print(k, ":", f[k])


def _parse_order_string(nr):
    '''
    解析报单字符串，形式可能为：10, 欠10，欠10报10，共，换10
    :param nr: 
    :return: 返回(ordered, owed)，分别为报单数量，欠货数量
    '''
    ordered = 0 # 报单数量
    owed = 0 # 欠货数量
    m = re.match(r'^([-+]?\d+)$', nr)
    if m != None:
        ordered = int(m.group(1))

    m = re.match(r'.*报(\d+).*', nr)
    if m != None:
        ordered = int(m.group(1))

    m = re.match(r'.*欠(\d+).*', nr)
    if m != None:
        owed = int(m.group(1))

    m = re.match(r'.*换(\d+).*', nr)
    if m != None:
        owed = owed + int(m.group(1)) # 注意此处为累加
    return (ordered, owed)

def _parse_payed_received_string(s):
    '''
    解析实付和实拿字符串。
    数字表示数量
    x,X表示0
    None表示空
    其他表示无异常
    :param nr: 
    :return: 数字：实际数量，None： 空， 'OK'：无异常 
    '''
    r = 0
    if s == None:
        return None

    m = re.match(r'^([-+]?\d+)$', s)
    if m != None:
        r = int(m.group(1))
        return r

    if s == 'x' or s == 'X':
        return 0


    return 'OK'

def calc_received_exceptions(nr_s, payed_s, received_s):
    '''
    计算到货异常，目前主要是欠货。
    
    :param nr: 
    :param payed: 实付，可能为空，数字，X，或其他
    :param received: 实拿，可能为空，数字，X，或其他
    :return: （totoal, owed）,total:总计应拿，owed：欠货，正数 - 档口欠我们，0-平衡，负数：我们欠档口
    '''

    ordered, owed = _parse_order_string(nr_s)
    payed = _parse_payed_received_string(payed_s)
    received = _parse_payed_received_string(received_s)
    total = ordered + owed

    if payed == None:
        payed = 0
    if payed == 'OK':
        payed = ordered

    if received == 'OK':
        return (total, 0) # 实拿为空或者没有问题，则无异常
    elif owed == 0 and received == None:
        return (total, 0) # 无欠货的情况下， 为空表示异常
    else:
        # 其余情况recevied为空视为0
        if received == None:
            received = 0
        balance = owed + payed - received
        return (total, balance)

if __name__ == "__main__":
    s = "\\"
    print(_parse_payed_received_string(s))