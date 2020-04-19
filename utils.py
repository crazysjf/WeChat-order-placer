# -*- coding: utf-8 -*-
from datetime import date
import re
import shutil
import os
import pandas as pd


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
        text = text + u" - 请微信拍照发单，2小时内付款。\n"
        text = text + u" - 单据请和货一起放入包裹。\n"
        text = text + u" - 有收版收色的情况，请第一时间反馈。\n"
        text = text + u" - 该单仅当日有效。合作愉快！\n"
        text = text + u" - 请尽量送拼包。\n\n"

        text = text + u" - 拼包点： 新金马后停车场，金富丽旁边，李艳 收，电话：18609665923，包裹上请写明“史小姐”和档口号，签字并拍照。\n\n"

        text = text + u" - 不能送拼包请联系工仔(15986644447，小郑)或报单微信。\n"





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
            e_str = "%s: %s, %s, "  % (p, l['code'], l['spec']) + gen_text_for_one_exception_line(l)
            print(e_str)
        print("-------------------")
    print("档口数： % s，异常数： % s\n" % (len(r.keys()), e_cnt))


def gen_all_orders_text(orders):
    s = u''
    for p in orders.keys():
        s = s + gen_order_text(orders, p)
        s = s + '\n'
    return s

def gen_text_for_one_exception_line(l, simplified=False):
    '''
    生成单行异常文本。格式：
    无异常标记：
    到x件欠x件
    
    有异常标记：
    备注
    
    注意文本里面不应该包含档口名称
    '''

    if not 'notation' in l.keys():
        if simplified == False:
            s = "\t到%s件，欠%s件" % (l['received'], l['nr'])
        else:
            s = "\t欠%s" % l['nr']
    else:
        s = l['notation']

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

            #s = s + "%-5s,\t%-10s,\t到%s件，欠%s件\n" % (code, l['spec'], l['received'], l['nr'])
            s = s + "%-5s,\t%-10s," % (code, l['spec']) + gen_text_for_one_exception_line(l) + '\n'
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
    if nr == None:
        return (0,0)
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
    #total = ordered + owed

    if payed == None:
        payed = 0
    if payed == 'OK':
        payed = ordered

    if received == 'OK' or received == None:
        return (received, 0) # 实拿没有问题，或者为空，则无异常
    # elif owed == 0 and received == None:
    #     return (total, 0) # 无欠货的情况下， 为空表示异常
    else:
        balance = owed + payed - received
        return (received, balance)

def backup_file(f):
    #m = re.match(r'(.*)\.xlsx', f)
    m = re.match(r'(.*)(\.[^.]*)', f)
    if m != None:
        f2 = m.group(1) + '-备份' + m.group(2)
    else:
        print('文件名格式错误')
        return
    if os.path.exists((f2)):
        i = input("备份文件：%s已存在。是否覆盖？(y/N)")
        if i != 'y':
            return

    shutil.copy2(f, f2)

def get_good_profile_file(tof):
    '''
    获取商品资料文件名。
    tof是当天计划采购建议报表，商品资料文件是同文件夹下的名为：商品资料xxxx.xlsx的文件。
    存在的话返回绝对路径，否则返回None。
    '''
    tof = os.path.abspath(tof)
    top = os.path.dirname(tof)
    #print(top)
    list = os.listdir(top)
    for f in list:
        #print(f)
        m = re.match(r'^商品资料.*xlsx$', f)
        if m != None:
            return os.path.join(top, m.group(0))
    return None

cols_to_reserve = ["款式编码", "商品编码", "供应商", "供应商款号", "颜色规格", "商品简称", "前7天<br/>销量", \
                   "待发货数","仓库库存数","建议采购数","商品备注","最早付款时间", "上限天数","成本价" ]

def analyze_annotaion(anno):
    """可能备注形式：
    **报11356-犇犇.8161
    **报11356-犇犇.8161*24
    **报11356-犇犇.8161, 其他备注
    **报11356-犇犇.8161.黄L*24"""

    m = re.match(r'.*\*\*报([^,*]+)[\*([0-9]+]',anno)
    if m is None:
        return

    str = m.group(1)
    print( m.group(2))

    print(str)

def process_xls(today_order_file):
    '''处理聚水潭导出报表。代替原来VBA代码'''
    df = pd.read_excel(today_order_file)
    columns = df.columns.tolist()
    for col in columns:
        if col not in cols_to_reserve:
            df.pop(col)

    #print(df['商品备注'])
    idx = df['商品备注'].apply(lambda s: "收" in str(s) or  "清" in str(s)  or "销低" in str(s))

    # 不报单商品：收清销低商品
    df_no_place = df.loc[idx]

    #print(df.loc[idx,["供应商", "商品编码", "供应商款号", "颜色规格", "前7天<br/>销量", "待发货数","仓库库存数","建议采购数","商品备注"]])

    df = df.loc[~idx]

    # 处理**报
    # 插入异常
    # 计算天数


    #输出
    out_file = os.path.join(os.path.dirname(today_order_file), "备份-" + os.path.basename(today_order_file))
    df.to_excel(out_file, index=False)



if __name__ == "__main__":
    s = "收4.1，**报11396-梦梦.5890*21"
    analyze_annotaion(s)

