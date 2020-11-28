from xls_processor import XlsProcessor
import itchat
import utils
import time
import config
import constants

def place_order(order_file, reverse=False):
    '''
    报单
    :param order_file: 报单报表 
    :return: 
    '''
    xls_processor = XlsProcessor(order_file)
    orders = xls_processor.gen_orders()
    if itchat.check_login() != "200":
        itchat.auto_login(hotReload=True, enableCmdQR=True)

    friends = itchat.get_friends()

    unknown_providers = []
    for p in orders.keys():

        f = utils.get_store(friends, p)
        if f == None:
            unknown_providers.append(p)

    print(u'报单内容：')
    print(utils.gen_all_orders_text(orders))
    print(u'\n共 %d 家\n' % len(orders.keys()))

    print(u'以下供应商未找到：')
    for p in unknown_providers:
        print(p)
    print('\n共 %d 家\n' % len(unknown_providers))


    while(True):
        #TODO: Continue这行只能用英文，用中文或者unicode会导致powershell中执行异常
        str = input("Continue? (y/N)")
        if str == 'y' or str == 'Y':
            for p in sorted(orders.keys(), reverse=reverse):
                f = utils.get_store(friends, p)
                if f != None:


                    order_text = utils.gen_order_text(orders, p)
                    itchat.send(order_text, toUserName=f['UserName'])
                    print(u"已发送：" + p)
                    time.sleep(constants.SENDING_TIME_INTERVAL)
            break
        elif str == 'n' or str == "N":
            break

    ret = xls_processor.annotate_unknown_providers(unknown_providers)


def send_today_exceptions(today_order_file):
    '''
    发送当天到货异常
    
    :param e: 异常字典
    :return: 
    '''
    if itchat.check_login() != "200":
        itchat.auto_login(hotReload=True, enableCmdQR=True)

    friends = itchat.get_friends()

    to = XlsProcessor(today_order_file)
    e = to.calc_order_exceptions()

    unknown_providers = []
    for p in e.keys():
        f = utils.get_store(friends, p)
        if f == None:
            unknown_providers.append(p)

    print('发送内容：')
    print(utils.gen_all_exception_text(e))
    print('\n共 %d 家\n' % len(e.keys()))

    print('以下供应商未找到：')
    for p in unknown_providers:
        print(p)
    print('\n共 %d 家\n' % len(unknown_providers))

    while(True):
        #TODO: Continue这行只能用英文，用中文或者unicode会导致powershell中执行异常
        str = input("Continue? (y/N)")
        if str == 'y' or str == 'Y':
            for p in e.keys():
                f = utils.get_store(friends, p)
                if f != None:
                    order_text = utils.gen_exception_text(e, p)
                    itchat.send(order_text, toUserName=f['UserName'])
                    print(u"已发送：" + p)
                    # 加入间隔，以免微信报错 ：发送消息太频繁。
                    time.sleep(constants.SENDING_TIME_INTERVAL)
            break
        elif str == 'n' or str == "N":
            break

def send_order_file(today_order_file):
    '''
    发送报表给采购
    :param today_order_file: 
    :return: 
    '''
    if itchat.check_login() != "200":
        itchat.auto_login(hotReload=True, enableCmdQR=True)

    friends = itchat.get_friends()

    purchaser = config.puerchaser_nickname
    f = utils.get_store(friends, purchaser)
    itchat.send_file(today_order_file, toUserName=f['UserName'])
    print("已发送：%s 至 %s" % (today_order_file, purchaser))