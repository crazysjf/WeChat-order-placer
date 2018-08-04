# -*- coding: utf-8 -*-
from openpyxl import load_workbook
import openpyxl.styles as sty
import utils

class Singleton(object):
    _instance = None
    def __new__(cls, *args, **kw):
        if not cls._instance:
            #cls._instance = super(Singleton, cls).__new__(cls, *args, **kw) # Python2
            cls._instance = object.__new__(cls)  # python3
        return cls._instance

class XlsProcessor(Singleton):
    def __init__(self, f):
        self._f = f

    # def _get_xls_column_nr(self, table_head, name):
    #     for i in range(0, len(table_head)):
    #         if table_head[i] == name:
    #             return i
    #     return None

    def _get_column_cn(self, ws, name):
        for c in range(1, ws.max_column):
            v = ws.cell(row = 1, column=c).value
            if v == name:
                return c
        return None

    def gen_orders(self):
        '''
        生成所有订单并返回，格式：
        {
        档口1: [{code: 商品编码1， spec: 规格1, nr: 数量1}，
                {code: 商品编码2， spec: 规格2, nr: 数量2，... ],
        档口2: [{code: 商品编码1， spec: 规格1, nr: 数量1}，
                {code: 商品编码2， spec: 规格2, nr: 数量2，... ],
        ...
        }
        '''
        wb = load_workbook(self._f)
        ws = wb.active
        nrows = ws.max_row

        provider_cn = self._get_column_cn(ws, u'供应商')
        code_cn = self._get_column_cn(ws, u'供应商款号')
        spec_cn = self._get_column_cn(ws, u'颜色规格')
        nr_cn = self._get_column_cn(ws, u'数量')

        orders = {}  # 解析之后的所有订单，键值为档口名
        provider_order = []  # 单个档口的订单

        for i in range(2, nrows):
            provider = utils.convert_possible_num_to_str(ws.cell(row=i, column=provider_cn).value)
            code = utils.convert_possible_num_to_str(ws.cell(row=i, column=code_cn).value)

            if provider == "**" or provider == u"样衣":
                # 截止到内容未**，或者“样衣”两个字的行
                break

            # 跳过空行和汇总行。
            if provider == None or provider.find(u'汇总') != -1:
                continue
            #print provider, code

            order_line = {}
            order_line['code'] = code
            order_line['spec'] = ws.cell(row=i, column=spec_cn).value
            order_line['nr'] = utils.convert_possible_num_to_str(ws.cell(row=i, column=nr_cn).value)

            provider_order.append(order_line)

            if provider in orders:
                orders[provider].append(order_line)
            else:
                orders[provider] = [order_line]

        wb.close()
        return orders

    def annotate_unknown_providers(self, unknown_provider):
        '''
        把未知供应商标记为红色。
        成功返回True
        如果文件被其他应用打开，标注失败，返回False。
        :param unknown_provider: 
        :return: 
        '''
        wb = load_workbook(self._f)
        ws = wb.active

        provider_cn = self._get_column_cn(ws, u'供应商')
        code_cn = self._get_column_cn(ws, u'供应商款号')

        # 写入数据
        for i in range(2, ws.max_row):
            code = utils.convert_possible_num_to_str(ws.cell(row=i, column=code_cn).value)
            # 跳过汇总行
            if code == None:
                continue

            cell = ws.cell(row=i, column=provider_cn)
            v = cell.value
            for p in unknown_provider:
                if v == p:
                    cell.fill = sty.PatternFill(fill_type='solid', fgColor="ff6347")
        try:
            wb.save(self._f)
        except IOError:
            return False

        wb.close()
        return True

if __name__ == "__main__":
    xp = XlsProcessor('./test.xlsx')
    orders = xp.gen_orders()
    import  utils
    print(utils.gen_all_orders_text(orders))

    up = [u'53d123', u'孙劲飞']
    print(xp.annotate_unknown_providers(up))
