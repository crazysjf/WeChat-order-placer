# -*- coding: utf-8 -*-
import itchat
import re

itchat.auto_login(hotReload=True)

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
        #print f
        #print remarkName
        m = re.match(pat, remarkName, re.IGNORECASE)
        if m != None:
            return f

    return None


#itchat.send('Hello, filehelper', toUserName='filehelper')
def print_friend(f):
    if f == None:
        print '无此好友'
        return
    for k in f.keys():
        print k, ":", f[k]

#r = itchat.search_friends()
#r = itchat.search_friends(nickName=u'14b452')
#r = itchat.get_friends()
#r = itchat.search_friends(remarkName=u'14b425-火爆服饰4B425')


friends = itchat.get_friends()

f = get_store(friends, u'13C292-雅雅')
print_friend(f)
#