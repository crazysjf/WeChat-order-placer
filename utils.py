# -*- coding: utf-8 -*-
from datetime import date, datetime
import re
import shutil
import os, os.path
import pandas as pd
import xls_processor
import math
import constants
import config

# 数字转字母
# 0=>A, 1=>B, 2=>C，以此类推
def num_to_alphabet(nr):
    return chr(ord('A') + nr)

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

        text = text + u" - 拼包点： 新金马后停车场，025档前，金富丽旁边，李艳 收，电话：18609665923，包裹上请写明“史小姐”和档口号，签字并拍照。\n\n"

        text = text + u" - 不能送拼包请联系工仔(15986644447，小郑)或报单微信。\n"


        text = text + u"\n\n - 大家都很忙，麻烦各位开实数单，不要欠货，减少对账麻烦。\n"
        text = text + u" - 欠货会导致第二天不报单，影响彼此业务。\n"
        text = text + u" - 感谢开实数单的各位的支持。\n"
        #text = text + u"\n 仓库忙不过来，今天是简化报单，明天恢复正常。\n"





    except:
        text = '报单生成错误'
    return text

def gen_exception_summary(r, only_positive=False):
    '''
    打印到货异常汇总。
    :param r: 异常字典
    :return:
    '''
    s = "\n拿货异常：\n"
    s = s + "-------------------\n"

    e_cnt = 0
    for p in r.keys():
        lines = r[p]
        for l in lines:
            if only_positive == True:
                if not 'notation' in l.keys() and l['nr'] < 0:
                    continue

            e_cnt = e_cnt + 1
            e_str = "%s: %s, %s, "  % (p, l['code'], l['spec']) + gen_text_for_one_exception_line(l)
            s = s + e_str + "\n"
        s = s + "\n-------------------\n\n"
    s = s + "档口数： % s，异常数： % s\n\n" % (len(r.keys()), e_cnt)
    return s

def print_exception_summary(r, only_positive=False):
    # '''
    # 打印到货异常汇总。
    # :param r: 异常字典
    # :return:
    # '''
    # print("\n拿货异常：")
    # print("-------------------")
    #
    # e_cnt = 0
    # for p in r.keys():
    #     lines = r[p]
    #     for l in lines:
    #         e_cnt = e_cnt + 1
    #         e_str = "%s: %s, %s, "  % (p, l['code'], l['spec']) + gen_text_for_one_exception_line(l)
    #         print(e_str)
    #     print("-------------------")
    # print("档口数： % s，异常数： % s\n" % (len(r.keys()), e_cnt))
    print(gen_exception_summary(r, only_positive))

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
    :return: （received, owed）,total:总计应拿，owed：欠货，正数 - 档口欠我们，0-平衡，负数：我们欠档口
    '''

    ordered, owed = _parse_order_string(nr_s)
    payed = _parse_payed_received_string(payed_s)
    received = _parse_payed_received_string(received_s)
    total = ordered + owed

    if payed == None:
        payed = 0
        #payed = ordered
    if payed == 'OK':
        payed = ordered

    if received == None:
        received = 0 # 收到为空即为0

    if received == 'OK':
        received = total

    # 实拿没有问题，或者为空，则无异常
    # elif owed == 0 and received == None:
    #     return (total, 0) # 无欠货的情况下， 为空表示异常
    #else:
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
    测试用例：
    str = "**报111280-美雪.11832-1.黄L*23,"
    str = "收4。3，**报111280-美雪.11832-1*23,"
    str = "收4。3，**报111280-美雪.11832-1"
    str = "收4。3，**报111280-美雪.11832-1.黄L"

    如果没有解析到内容，返回None。
    如果解析到内容，则返回字典，格式如下：
    {
        provider:
        code:
        spec:
        price:
    }
    如果对应键没有内容，则值为None
    """
    m = re.match(
        r'.*\*\*报(?P<provider>[^.,*]+)\.(?P<code>[^.,*]+)(?:\.(?P<spec>[^.,*]+)){0,1}(?:\*(?P<price>[0-9]+)){0,1}.*',
        anno)

    if m is None:
        return None

    ret = {}
    ret['provider'] = m.group('provider')
    ret['code'] = m.group('code')
    ret['spec'] = m.group('spec')
    ret['price'] = m.group('price')
    return ret


#  计算报单数量
def calNum(l):
    n = l['建议采购数']
    upper_lim = l['上限天数']

    if upper_lim != upper_lim: # 上限为nan, 不压货
        if n <= 20:
            ret = math.ceil(n/5)*5
        else:
            ret = math.ceil(n/10)*10
    else: # 压货
        if n <= 7:
            ret = 5
        elif 7 < n and n <= 12:
            ret = 10
        elif 12 < n and n <= 17:
            ret = 15
        elif 17 < n and n <= 24:
            ret = 20
        else:
            ret = math.floor(n/10)*10

    return ret

def gen_defectives_data(yesterday_defective_file, goods_file):
    """
    生成次品数据。

    :param yesterday_defective_file: 昨日扫描生成的次品登记文件
    :param goods_file:  普通商品资料导出文件
    :return: (defe_df, ignored_defe_good_df),
    defe_df: 次品数据
    形式：
    款式编码  商品编码    供应商         供应商商品款号     数量      价格  备注
    xxx       xxx         xxx           xxx                xxx      xxx   xxx

    忽略尺码和颜色，返回数量是同款各个数量之和

    ignored_defe_good_df：根据需要忽略掉的供应商次品数据。

    """
    if os.path.exists(yesterday_defective_file):
        df = pd.read_excel(yesterday_defective_file)
        print("正在解析商品资料...")
        goods_df = pd.read_excel(goods_file)
        print("完成")
    else:
        return None, None

    df = pd.merge(df, goods_df, how='left', left_on="商品编码", right_on="商品编码")

    # merge之后的表头：
    # 图片     退货单号       仓库       供应商  采购单号               单据日期   状态  备注_x  物流公司  物流单号
    # 商品编码  商品名称 颜色及规格_x  数量  单价  成本价比例  基本售价_x  金额  财审人  财审日期  备注1        款式编号  供应商货号  创建人  标记多标签
    # 图片地址        款式编码        国际条形码   商品名  商品简称 商品属性  单位 颜色及规格_y  基本售价_y  市场|吊牌价
    # 成本价  其它价格1  其它价格2  其它价格3  重量   仓位   分类  虚拟分类       供应商名      供应商编号  供应商商品编码
    # 供应商商品款号  品牌 备注_y                创建时间                修改时间  库存同步  自动上架


    pd.options.display.max_rows = 1000
    pd.options.display.max_columns = 100
    pd.options.display.width = 300

    # 无供应商警告
    for rIndex in df.index:
        l = df.loc[rIndex]
        p = l['供应商名']
        if p != p: # p为 Nan
            print("%s无供应商，需要更新商品资料" % l['商品编码'])

            # 给为Nan的供应商名和供应商款号赋值，避免在后面的drop_duplicates时被删除
            df.loc[rIndex, "供应商名"] = l['商品编码']
            df.loc[rIndex, "供应商商品款号"] = l['商品编码']

    # 过滤滤指定供应商
    ignored_defe_df = df[df["供应商名"].apply(lambda p: p in constants.PROVIDERS_IGNORING_DEFECTIVE_GOODS)]
    df = df[df["供应商名"].apply(lambda p: p not in constants.PROVIDERS_IGNORING_DEFECTIVE_GOODS)]


    ret_df = df.drop_duplicates(subset=['供应商名', '供应商商品款号'], keep='first')
    ret_df = ret_df.copy() # 不加这行会出现set on a copy of a slice from a DataFrame warning。具体原因不明

    for r in ret_df.index:
        provider = ret_df.loc[r]["供应商名"]
        code = ret_df.loc[r]["供应商商品款号"]
        tmp_df = df[(df['供应商名'] == provider) & (df['供应商商品款号'] == code)]
        sum = tmp_df['退货数量'].apply(lambda x: int(x)).sum() # 求和
        ret_df.loc[r, '退货数量'] = sum

    def _process(df):
        df = df[['款式编码','商品编码', '供应商名', "供应商商品款号", "退货数量", "成本价", "备注"]]
        df.rename(columns={"供应商名":"供应商", "供应商商品款号":"供应商款号", "退货数量":"数量", "备注":"商品备注", }, inplace=True)
        return df

    (ret_df, ignored_defe_df) = map(_process, (ret_df, ignored_defe_df))
    # ret = ret_df[['款式编码','商品编码', '供应商名', "供应商商品款号", "数量", "成本价", "备注_y"]]
    # ret.rename(columns={"供应商名":"供应商", "备注_y":"商品备注", "供应商商品款号":"供应商款号"}, inplace=True)
    return (ret_df, ignored_defe_df)


def process_xls(today_order_file, yestoday_order_file, yesterday_defective_file=None, goods_file=None):
    '''处理聚水潭导出报表。代替原来VBA代码'''
    # test
    # pd.options.display.max_columns = 100
    # pd.options.display.width = 300

    df = pd.read_excel(today_order_file)

    # 删除不需要列
    columns = df.columns.tolist()
    for col in columns:
        if col not in cols_to_reserve:
            df.pop(col)

    # 插入一些列
    c = df.columns.get_loc('商品备注')
    df.insert(c, "实拿", "")  # 注意由于c不变，插入后顺序和这里相反
    df.insert(c, "实付", "")
    df.insert(c, "数量", "")

    # 计算数量列
    num = df.apply(calNum, axis=1)
    df['数量'] = num

    # 处理次品
    if yesterday_defective_file is not None:
        (defe_df, ignored_df) = gen_defectives_data(yesterday_defective_file, goods_file)
        if defe_df is not None:
            defe_df['数量'] = defe_df['数量'].apply(lambda n: '换' + str(n))  # 次品数量前面 + 换
            defe_df['供应商款号'] = defe_df['供应商款号'].apply(lambda n: str(n) + "(次)")  # 次品编码后面 + 次
            defe_df['颜色规格'] = '次品' + date.today().strftime("-%m.%d")

            df = pd.concat([df, defe_df], axis=0, sort=False)
            df.reset_index(inplace=True) # 连接后索引必须重置

        if ignored_df is not None and len(ignored_df) > 0:
            f = os.path.dirname(today_order_file) + "\未上报次品.xlsx"
            ignored_df.to_excel(f, index=False)

    # 处理**报
    for ridx in df.index:
        # 次品忽略**报
        if "次" in str(df.loc[ridx]['供应商款号']):
            continue

        notation = df.loc[ridx]['商品备注']
        if notation == notation: # nan判断: nan!=nan, 备注为空的时候，这里值为nan
            ret = analyze_annotaion(notation)
            if ret is not None:
                #print(ret)
                if ret['provider'] is not None:
                    df.loc[ridx, '供应商'] = ret['provider']
                if ret['code'] is not None:
                    df.loc[ridx, '供应商款号'] = ret['code']
                if ret['spec'] is not None:
                    df.loc[ridx, '颜色规格'] = ret['spec']
                if ret['price'] is not None:
                    df.loc[ridx, '成本价'] = float(ret['price'])


    # 插入异常
    if yestoday_order_file is None:
        input("未指定前日报表文件。按任意键继续...")
    yo = xls_processor.XlsProcessor(yestoday_order_file)
    e = yo.calc_order_exceptions()

    for p in e.keys():
        for l in e[p]:

            # 欠货如果为负数则不插入
            if not 'notation' in l.keys() and  l['nr'] < 0:
                continue


            c = l['code']
            s = l['spec']

            idx = (df['供应商'] == p) & (df['供应商款号'] == c) & (df['颜色规格'] == s)
            list = df[idx].index.tolist()
            text = gen_text_for_one_exception_line(l, True)
            #print(text)
            if len(list) == 0:  # 表格中没有对应行
                df = df.append({"供应商":p, "供应商款号":c, "颜色规格": s, "数量":text}, ignore_index = True)
            elif len(list) == 1: # 表格中找到唯一对应行
                # 如果欠货少于报单数，要写成“欠x报y，共z”的形式
                i = list[0]
                if not 'notation' in l.keys():
                   suggested_nr = df.loc[i, "建议采购数"]
                   e_nr = l['nr']

                   if e_nr < suggested_nr and config.whether_place_order_when_owed_goods is True:
                       order_nr = (int((suggested_nr - e_nr) / 5) + 1) * 5  # 向上取整到5
                       v = text + "报%d，共%d" % (order_nr, e_nr + order_nr)
                       df.loc[i, "数量"] = v
                   else:
                       df.loc[i, "数量"] = text
                else:
                    v = text + ',' + str(df.loc[list[0]]["数量"])
                    df.loc[i, "数量"] = v

            else: # 报表有异常，有重复行
                print("报表异常，多个位置发现同样商品：", p, l, list)


    # 计算天数
    def cal_days(s):
        try:
            date = datetime.strptime(s, "%Y/%m/%d %H:%M:%S")
            now = datetime.now()
            d = ((now - date).total_seconds()/60/60 + 12) / 24
            d = int(d)
        except:
            d = 0
        return "D%s" % d

    d = df['最早付款时间'].apply(cal_days)
    df.insert(df.columns.get_loc('最早付款时间'),"天数",d )
    df.pop('最早付款时间')


    # 更新备注
    def update_annotation(l):
        code = l['商品编码']
        anno = l['商品备注']
        if anno != anno:
            anno = ""

        if code == code: # 商品编码不为空
            #(p, c, s, *_) = code.split('-')
            # 不要用tupple unpack以防编码格式不对导致异常
            splited = code.split('-')
            p = splited[0]
            c = splited[1] if len(splited) > 1 else ""
            s = splited[2] if len(splited) > 2 else ""

            if     not p.upper() in str(l['供应商']).upper() or \
                    not c.upper() in str(l['供应商款号']).upper() or \
                    not s.upper() in str(l['颜色规格']).upper():

                #print(code, l['供应商'],l['供应商款号'],l['颜色规格'])
                #print(anno, code)
                if len(anno) == 0:
                    return code
                else:
                    return anno + '\n' + code
            else:
                return anno
        else:
            return anno

    s = df.apply(update_annotation, axis=1)
    df['商品备注'] = s


    # 排序
    # 供应商转大写后排序，然后删除，保留原供应商大小写，以免异常插入出问题
    df['供应商大写'] = df['供应商'].apply(lambda p: str(p).upper())

    # 对一个档口，档口次品排所有商品之后
    def is_defe(x):
        if '次' in str(x):
            return 1
        else:
            return 0
    df['是否次品'] = df["供应商款号"].apply(is_defe)

    # 尺码按常识排序
    def size_to_num(x):
        x = str(x).upper()
        map = (('XXXS', '1'),  ('XXXL','9'),
               ('XXS','2'), ('XXL','8'),
               ('XS','3'), ('XL','7'),
               ('S','4'),('M','5'),('L','6'))
        for t in map:
            s = t[0]
            num = t[1]
            if s in x:
                return x.replace(s,num)
        return x

    df['尺码改数字'] = df["颜色规格"].apply(size_to_num)



    df = df.sort_values(["供应商大写","是否次品", "供应商款号","尺码改数字"])
    del df['供应商大写']
    del df['是否次品']
    del df['尺码改数字']


    # 删除不报单商品：收清销低商品 ,但不包含有异常的款
    def lam(l):
        s = str(l['商品备注'])
        nr = str(l['数量'])

        p = str(l['供应商'])
        code = str(l['供应商款号'])
        spec = str(l['颜色规格'])

        in_exceptions = False # 是否处于异常中
        if p in e.keys():
            for ll in e[p]:
                if ll['code'] == code and ll['spec'] == spec:
                    in_exceptions = True

        return ("收" in s or "清" in s or "销低" in s) and (in_exceptions is not True) and ("次" not in spec)#"欠" not in nr  and "换" not in nr

    #idx = df['商品备注'].apply(lambda s: ("收" in str(s) or "清" in str(s) or "销低" in str(s)) and "欠" not in str(s) and "换" not in str(s))
    idx = df.apply(lam, axis=1)

    df_no_place = df.loc[idx]
    df = df.loc[~idx]

    # 把收清商品添到最后
    df.reset_index(inplace=True, drop=True)
    df_no_place.reset_index(inplace=True, drop=True)
    df = df.append(df_no_place, ignore_index=True)

    # 输出
    if constants.TEST == True:
        f = os.path.dirname(today_order_file) + "\测试结果.xlsx"
        df.to_excel(f, index=False)
        xls_processor.XlsProcessor(f).format()
    else:
        backup_file(today_order_file)
        df.to_excel(today_order_file, index=False)
        xls_processor.XlsProcessor(today_order_file).format()


if __name__ == "__main__":
    s = "收4.1，**报11396-梦梦.5890*21"
    analyze_annotaion(s)

