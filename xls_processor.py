# -*- coding: utf-8 -*-
import xlrd

class Singleton(object):
    _instance = None
    def __new__(cls, *args, **kw):
        if not cls._instance:
            cls._instance = super(Singleton, cls).__new__(cls, *args, **kw)
        return cls._instance

class XlsProcessor(Singleton):
    def __init__(self, f):
        self._f = f

    def _get_xls_column_nr(self, table_head, name):
        for i in range(0, len(table_head)):
            if table_head[i] == name:
                return i
        return None

    def _convert_possible_float_to_str(self, v):
        '如果v是浮点，则转为字符串，否则原样返回'
        if type(v).__name__ == 'float':
            return str(int(v))
        else:
            return v

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
        sheet = xlrd.open_workbook(self._f).sheets()[0]
        nrows = sheet.nrows
        head = sheet.row_values(0)
        self.provider_cn = self._get_xls_column_nr(head, u'供应商')
        self.code_cn = self._get_xls_column_nr(head, u'供应商款号')
        self.spec_cn = self._get_xls_column_nr(head, u'颜色规格')
        self.nr_cn = self._get_xls_column_nr(head, u'数量')

        orders = {}  # 解析之后的所有订单，键值为档口名
        provider_order = []  # 单个档口的订单

        old_provider = None
        for i in range(1, nrows):
            provider = self._convert_possible_float_to_str(sheet.cell(i, self.provider_cn).value)

            if provider == "":
                continue
            if provider == "**":
                break

            order_line = {}
            order_line['code'] = self._convert_possible_float_to_str(sheet.cell(i, self.code_cn).value)
            order_line['spec'] = sheet.cell(i, self.spec_cn).value
            order_line['nr'] = self._convert_possible_float_to_str(sheet.cell(i, self.nr_cn).value)

            provider_order.append(order_line)

            if orders.has_key(provider):
                orders[provider].append(order_line)
            else:
                orders[provider] = [order_line]

        return orders