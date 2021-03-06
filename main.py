# -*- coding: utf-8 -*-
import utils
import sys
import getopt
import business_logic
import xls_processor
from goods_profile import GoodsProfile
from good_op_log import GoodsOpLog
import os
import re
import constants

def usage():
    print('''test.py [options]
options:
 -h, --help         help
 -t <当日报表>,      指定当日报表文件
 -y <昨日报表>,      指定昨日报表文件
 
 如果没有指任何报表文件，则自动分析。''')

def help():
    print(
"""命令帮助：
 h: 显示该帮助
 p: 报单
 ye: 显示昨天到货异常(yestoday exceptions)
 iye: 在当天报表里面插入昨天到货异常(insert yestoday exceptions)
 te: 显示今天到货异常(today exceptions)
 ste: 发送今天到货异常(send today exceptions)
 sof: 发送报表给采购(send order file)
 x: 处理Excel报表
 i: 显示参数信息
 q: 退出
 """)

def print_file_info():
    print("\n")
    print("当日报表：", today_order_file)
    print("昨日报表：", yesterday_order_file)
    print("\n")
    print("次品登记：", yesterday_defective_file)
    print("商品资料：", goods_file)
    print("商品操作日志：", good_op_log_file)
    print("\n")

try:
    options,args = getopt.getopt(sys.argv[1:],"hy:t:")
except getopt.GetoptError:
    usage()
    sys.exit()

today_order_file = None
yesterday_order_file = None
yesterday_defective_file = None # 次品文件
goods_file = None # 商品资料文件
good_op_log_file = None # 商品操作日志


for name,value in options:
    if name in ("-h",):
        usage()
    if name in ("-t",):
        today_order_file = value
    if name in ("-y",):
        yesterday_order_file = value

if today_order_file == None and yesterday_order_file == None:
    print("参数不全")
    usage()
    exit()

# 在昨日报表文件夹中查找次品登记文件
if yesterday_order_file is not None:
    yof_dir = os.path.dirname(yesterday_order_file)
    fs = os.listdir(yof_dir)
    for f in fs:
        m = re.match(r'.*采购退货.*\.xlsx$', f)
        if m != None:
            yesterday_defective_file = os.path.join(yof_dir, f)

    # 向上面最多追溯4级查找商品资料文件
    dir = yof_dir
    for i in range(0, 4):
        dir = os.path.dirname(dir) # 路径向上走一级
        fs = os.listdir(dir)
        for f in fs:
            m = re.match('商品信息.*\.xlsx$', f)
            if m != None:
                goods_file = os.path.join(dir, f)
                break

dir = os.path.dirname(today_order_file)
fs = os.listdir(dir)
for f in fs:
    m = re.match('[^~]*商品操作日志.*\.xlsx$', f)
    if m != None:
        good_op_log_file = os.path.join(dir, f)
        break

print_file_info()

if constants.TEST == True:
    utils.process_xls(today_order_file, yesterday_order_file, yesterday_defective_file, goods_file)
    exit(0)

while True:
    cmd = input("输入命令(h：帮助)：")
    if cmd == "h":
        help()

    elif cmd == "p":
        business_logic.place_order(today_order_file)
        print("报单完成")
        business_logic.send_order_file(today_order_file)

    elif cmd == "rp": # 档口反序报单
        business_logic.place_order(today_order_file, reverse=True)
        print("报单完成")
        business_logic.send_order_file(today_order_file)

    elif cmd == "ye":
        yo = xls_processor.XlsProcessor(yesterday_order_file)
        r = yo.calc_order_exceptions()
        utils.print_exception_summary(r)

    elif cmd == "ype": # 仅显示欠货为正数的异常
        to = xls_processor.XlsProcessor(yesterday_order_file)
        r = to.calc_order_exceptions()
        utils.print_exception_summary(r, only_positive = True)

    elif cmd == "te":
        to = xls_processor.XlsProcessor(today_order_file)
        r = to.calc_order_exceptions()
        utils.print_exception_summary(r)

    elif cmd == "tpe": # 仅显示欠货为正数的异常
        to = xls_processor.XlsProcessor(today_order_file)
        r = to.calc_order_exceptions()
        utils.print_exception_summary(r, only_positive = True)

    elif cmd == "ste":
        business_logic.send_today_exceptions(today_order_file)

    elif cmd == "rte":
        # 刷新今日异常，直接在当前报表上操作
        xls_processor.XlsProcessor(today_order_file).refresh_today_exceptions(good_op_log_file)

    elif cmd == "gof":
        GoodsOpLog(good_op_log_file).test()

    elif cmd == "iye":
        c = input('iye命令已废除，确定要使用？(y/N)：')
        if c != 'y' and c != 'Y':
            continue

        yo = xls_processor.XlsProcessor(yesterday_order_file)
        e = yo.calc_order_exceptions()
        utils.print_exception_summary(e)
        utils.backup_file(today_order_file)
        to = xls_processor.XlsProcessor(today_order_file)
        to.insert_exceptions(e)

        # 处理无仓位货品

    elif cmd == "sof":
        business_logic.send_order_file(today_order_file)


    elif cmd == 'i':
        print_file_info()

    elif cmd == 'x':
        utils.process_xls(today_order_file, yesterday_order_file, yesterday_defective_file, goods_file)

    elif cmd == 'c':
        utils.gen_defectives_data(yesterday_defective_file, goods_file)

    elif cmd == "q":
        exit()

    else:
        help()


