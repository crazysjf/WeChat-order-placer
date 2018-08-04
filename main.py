# -*- coding: utf-8 -*-
import sys
import sys
import getopt
import business_logic

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

if today_order_file == None or yestoday_order_file == None:
    print("参数不全")
    usage()
    exit()

def help():
    print(
"""命令帮助：
 h: 显示该帮助
 p: 报单
 q: 退出
 """)


while True:
    cmd = input("输入命令(h：帮助)：")
    if cmd == "h":
        help()
    elif cmd == "p":
        business_logic.place_order(today_order_file)
        exit()
    elif cmd == "q":
        exit()
    else:
        help()


