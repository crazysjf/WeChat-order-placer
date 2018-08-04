from xls_processor import XlsProcessor
import itchat
import utils
import time

def place_order(order_file):
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
    print()
    print(u'共 %d 家' % len(orders.keys()))
    print()

    print(u'以下供应商未找到：')
    for p in unknown_providers:
        print(p)
    print()
    print(u'共 %d 家' % len(unknown_providers))
    print()

    while(True):
        #TODO: Continue这行只能用英文，用中文或者unicode会导致powershell中执行异常
        str = input("Continue? (y/N)")
        if str == 'y' or str == 'Y':
            for p in orders.keys():
                f = utils.get_store(friends, p)
                if f != None:
                    order_text = utils.gen_order_text(orders, p)
                    itchat.send(order_text, toUserName=f['UserName'])
                    print(u"已发送：" + p)
                    time.sleep(0.5) # 加入间隔，以免微信报错 ：发送消息太频繁。
            break
        elif str == 'n' or str == "N":
            break

    while(True):
        ret = xls_processor.annotate_unknown_providers(unknown_providers)
        if ret == True:
            break
        str = input("Annotaion failed. File may be open in other application. Retry? (Y/n)")
        if str == 'n' or str == 'N':
            break
