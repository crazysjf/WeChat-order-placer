# -*- coding: utf-8 -*-
import itchat
import re
import sys
import utils
from xls_processor import XlsProcessor

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
            print "e.message: %s" % code
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


def place_order():
    print "order placed"

if len(sys.argv) == 1:
    print u"Error： need xls file as parameter."
    exit(-1)
    #order_xls_file = r'D:\projects\20180421-WeChat-order-placer\src\4.21.xlsx'
    #order_xls_file = r'D:\projects\20180421-WeChat-order-plxxacer\src\testxx.xlsx'
else:
    order_xls_file = sys.argv[1]

xls_processor = XlsProcessor(order_xls_file)
orders = xls_processor.gen_orders()

itchat.auto_login(hotReload=True, enableCmdQR=True)

friends = itchat.get_friends()



unknown_providers = []
for p in orders.keys():

    f = get_store(friends, p)
    if f == None:
        unknown_providers.append(p)

print u'报单内容：'
print utils.gen_all_orders_text(orders)
print

print u'以下供应商未找到：'
for p in unknown_providers:
    print p
print


while(True):
    #TODO: Continue这行只能用英文，用中文或者unicode会导致powershell中执行异常
    str = raw_input("Continue? (y/N)")
    if str == 'y' or str == 'Y':
        for p in orders.keys():
            f = get_store(friends, p)
            if f != None:
                order_text = utils.gen_order_text(orders, p)
                itchat.send(order_text, toUserName=f['UserName'])
        break
    elif str == 'n' or str == "N":
        break

while(True):
    ret = xls_processor.annotate_unknown_providers(unknown_providers)
    if ret == True:
        break
    str = raw_input("Annotaion failed. File may be open in other application. Retry? (Y/n)")
    if str == 'n' or str == 'N':
        break
