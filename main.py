# -*- coding: utf-8 -*-
import itchat
import re
import xlrd
import sys
from datetime import date, datetime

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
        pat = r'^' + code
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
        print '无此好友'
        return
    for k in f.keys():
        print k, ":", f[k]

def get_xls_column_nr(table_head, name):
    for i in range(0, len(table_head)):
        if table_head[i] == name:
            return i
    return None

def convert_possible_float_to_str(v):
    '如果v是浮点，则转为字符串，否则原样返回'
    if type(v).__name__ == 'float':
        return str(int(v))
    else:
        return v

def gen_order_text(orders, p):
    '生成报单文本'
    text = ''
    try:
        t = date.today()  # 仅获取日期
        text = u'%s.%s 报单:\n\n' % (t.month, t.day)
        o = orders[p]
        for l in o:
            s = "%s: %-5s,\t%-10s,\t%-5s\n" % (p , l['code'], l['spec'] , l['nr'])
            text = text + s
    except:
        text = '报单生成错误'
    return text

def place_order():
    print "order placed"

if len(sys.argv) == 1:
    #order_xls_file = r'D:\projects\20180421-WeChat-order-placer\src\4.21.xlsx'
    order_xls_file = r'D:\projects\20180421-WeChat-order-plxxacer\src\testxx.xlsx'
else:
    order_xls_file = sys.argv[1]

sheet = xlrd.open_workbook(order_xls_file).sheets()[0]
nrows = sheet.nrows
head = sheet.row_values(0)
provider_cn = get_xls_column_nr(head, u'供应商')
code_cn = get_xls_column_nr(head, u'供应商款号')
spec_cn = get_xls_column_nr(head, u'颜色规格')
nr_cn   = get_xls_column_nr(head, u'数量')

orders = {} # 解析之后的所有订单，键值为档口名
provider_order = []  # 单个档口的订单

old_provider = None
for i in range(1, nrows):
    provider = convert_possible_float_to_str(sheet.cell(i, provider_cn).value)

    if provider == "":
        continue
    if provider == "**":
        break

    order_line = {}
    order_line['code'] = convert_possible_float_to_str(sheet.cell(i, code_cn).value)
    order_line['spec'] = sheet.cell(i, spec_cn).value
    order_line['nr']   = convert_possible_float_to_str(sheet.cell(i, nr_cn).value)

    provider_order.append(order_line)

    if orders.has_key(provider):
        orders[provider].append(order_line)
    else:
        orders[provider] = [order_line]

def print_all_order_text(orders):
    for p in orders.keys():
        print gen_order_text(orders, p)
        print


itchat.auto_login(hotReload=True)

friends = itchat.get_friends()

unfound_provider = []
for p in orders.keys():
    f = get_store(friends, p)
    if f == None:
        unfound_provider.append(p)

print u'报单内容：'
print_all_order_text(orders)
print

print u'以下供应商未找到：'
for p in unfound_provider:
    print p
print

while(True):
    #TODO: Continue这行只能用英文，用中文或者unicode会导致powershell中执行异常
    str = raw_input("Continue? (y/N)")
    if str == 'y' or str == 'Y':
        for p in orders.keys():
            f = get_store(friends, p)
            if f != None:
                order_text = gen_order_text(orders, p)
                itchat.send(order_text, toUserName=f['UserName'])
        exit()
    elif str == 'n' or str == "N":
        exit()


