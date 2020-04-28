# -*- coding: utf-8 -*-
import utils
import sys
import getopt
import business_logic
import xls_processor
from goods_profile import GoodsProfile
import os


def usage():
    print('''test.py [options]
options:
 -h, --help         help
 -t <当日报表>,      指定当日报表文件
 -y <昨日报表>,      指定昨日报表文件
 
 如果没有指任何报表文件，则自动分析。''')

try:
    options,args = getopt.getopt(sys.argv[1:],"hy:t:")
except getopt.GetoptError:
    usage()
    sys.exit()

today_order_file = None
yestoday_order_file = None

for name,value in options:
    if name in ("-h",):
        usage()
    if name in ("-t",):
        today_order_file = value
    if name in ("-y",):
        yestoday_order_file = value

print("当日报表：", today_order_file)
print("昨日报表：", yestoday_order_file)

if today_order_file == None and yestoday_order_file == None:
    print("参数不全")
    usage()
    exit()

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


if utils.TEST == True:
    utils.process_xls(today_order_file, yestoday_order_file)
    exit(0)

while True:
    cmd = input("输入命令(h：帮助)：")
    if cmd == "h":
        help()

    elif cmd == "p":
        business_logic.place_order(today_order_file)
        print("报单完成")
        business_logic.send_order_file(today_order_file)

    elif cmd == "ye":
        yo = xls_processor.XlsProcessor(yestoday_order_file)
        r = yo.calc_order_exceptions()
        utils.print_exception_summary(r)

    elif cmd == "te":
        to = xls_processor.XlsProcessor(today_order_file)
        r = to.calc_order_exceptions()
        utils.print_exception_summary(r)

    elif cmd == "ste":
        business_logic.send_today_exceptions(today_order_file)

    elif cmd == "iye":
        c = input('iye命令已废除，确定要使用？(y/N)：')
        if c != 'y' and c != 'Y':
            continue

        yo = xls_processor.XlsProcessor(yestoday_order_file)
        e = yo.calc_order_exceptions()
        utils.print_exception_summary(e)
        utils.backup_file(today_order_file)
        to = xls_processor.XlsProcessor(today_order_file)
        to.insert_exceptions(e)

        # 处理无仓位货品

    elif cmd == "sof":
        business_logic.send_order_file(today_order_file)


    elif cmd == 'i':
        print("当日报表文件：%s" % today_order_file)
        print("昨日报表文件：%s" % yestoday_order_file)

    elif cmd == 'x':
        utils.process_xls(today_order_file, yestoday_order_file)

    elif cmd == "q":
        exit()

    else:
        help()


