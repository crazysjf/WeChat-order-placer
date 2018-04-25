# -*- coding: utf-8 -*-
import xlrd, xlwt
from xlutils.copy import copy

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
            if provider == "**" or provider == u"样衣":
                # 截止到内容未**，或者“样衣”两个字的行
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

    def _change_cell_style(self, sheet,  x, y):
        pattern = xlwt.Pattern()  # Create the Pattern
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern_fore_colour = 5  # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        style = xlwt.XFStyle()  # Create the Pattern
        style.pattern = pattern  # Add Pattern to Style
        sheet.write(x, y, 'Cell Contents', style)

    def annotate_unknown_provider(self, unknown_provider):
        old_excel = xlrd.open_workbook(self._f, formatting_info=True)
        new_excel = copy(old_excel)
        sheet = new_excel.get_sheet(0)

        # 写入数据
        for p in unknown_provider:
            for i in range(1, sheet.nrows):
                if sheet.cell(i, self.provider_cn).value == p:
                    self._change_cell_style(sheet, i, self.provider_cn)

        new_excel.save(self._f)