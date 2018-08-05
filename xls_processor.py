# -*- coding: utf-8 -*-
from openpyxl import load_workbook
import openpyxl.styles as sty
import utils
import re

class XlsProcessor():
    def __init__(self, f):
        self._f = f

    # def _get_xls_column_nr(self, table_head, name):
    #     for i in range(0, len(table_head)):
    #         if table_head[i] == name:
    #             return i
    #     return None

    def _open(self):
        self.wb = load_workbook(self._f)
        self.ws = self.wb.active
        self.nrows = self.ws.max_row

    def _close(self):
        self.wb.close()
        self.wb = None
        self.ws = None
        self.nrows = 0

    def _save(self):
        self.wb.save(self._f)

    def _get_column_cn(self, name):
        for c in range(1, self.ws.max_column):
            v = self.ws.cell(row = 1, column=c).value
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
        # wb = load_workbook(self._f)
        # ws = wb.active
        # nrows = ws.max_row

        self._open()

        provider_cn = self._get_column_cn( u'供应商')
        code_cn = self._get_column_cn(u'供应商款号')
        spec_cn = self._get_column_cn(u'颜色规格')
        nr_cn = self._get_column_cn(u'数量')

        orders = {}  # 解析之后的所有订单，键值为档口名
        provider_order = []  # 单个档口的订单

        for i in range(2, self.nrows):
            provider = utils.convert_possible_num_to_str(self.ws.cell(row=i, column=provider_cn).value)
            code = utils.convert_possible_num_to_str(self.ws.cell(row=i, column=code_cn).value)

            if provider == "**" or provider == u"样衣":
                # 截止到内容未**，或者“样衣”两个字的行
                break

            # 跳过空行和汇总行。
            if provider == None or \
                            provider.find(u'汇总') != -1 or \
                            provider.find(u'总计') != -1:
                continue
            #print provider, code

            order_line = {}
            order_line['code'] = code
            order_line['spec'] = self.ws.cell(row=i, column=spec_cn).value
            order_line['nr'] = utils.convert_possible_num_to_str(self.ws.cell(row=i, column=nr_cn).value)

            provider_order.append(order_line)

            if provider in orders:
                orders[provider].append(order_line)
            else:
                orders[provider] = [order_line]

        self._close()
        return orders

    def annotate_unknown_providers(self, unknown_provider):
        '''
        把未知供应商标记为红色。
        成功返回True
        如果文件被其他应用打开，标注失败，返回False。
        :param unknown_provider: 
        :return: 
        '''
        self._open()

        provider_cn = self._get_column_cn(u'供应商')
        code_cn = self._get_column_cn(u'供应商款号')

        # 写入数据
        for i in range(2, self.nrows):
            code = utils.convert_possible_num_to_str(self.ws.cell(row=i, column=code_cn).value)
            # 跳过汇总行
            if code == None:
                continue

            cell = self.ws.cell(row=i, column=provider_cn)
            v = cell.value
            for p in unknown_provider:
                if v == p:
                    cell.fill = sty.PatternFill(fill_type='solid', fgColor="ff6347")
        try:
            self._save()
        except IOError:
            return False

        self._close()
        return True


    def calc_order_exceptions(self):
        '''获取报单异常，包括：
        欠货，
        其他
        '''
        # 处理欠，报，换等几种情况
        self._open()
        provider_cn = self._get_column_cn(u'供应商')
        code_cn = self._get_column_cn(u'供应商款号')
        spec_cn = self._get_column_cn(u'颜色规格')
        nr_cn = self._get_column_cn(u'数量')
        payed_cn = self._get_column_cn("实付")
        received_cn = self._get_column_cn("实拿")

        result = {}

        for i in range(2, self.nrows):
            provider = utils.convert_possible_num_to_str(self.ws.cell(row=i, column=provider_cn).value)
            code = utils.convert_possible_num_to_str(self.ws.cell(row=i, column=code_cn).value)
            spec = self.ws.cell(row=i, column=spec_cn).value

            # 截止到内容未**，或者“样衣”两个字的行
            if provider == "**" or provider == u"样衣":
                break

            # 跳过空行和汇总行。
            if provider == None or \
                            provider.find(u'汇总') != -1 or \
                            provider.find(u'总计') != -1:
                continue

            nr =  utils.convert_possible_num_to_str(self.ws.cell(row=i, column=nr_cn).value)
            payed = utils.convert_possible_num_to_str(self.ws.cell(row=i, column=payed_cn).value)
            received = utils.convert_possible_num_to_str(self.ws.cell(row=i, column=received_cn).value)
            total, e = utils.calc_received_exceptions(nr, payed, received)
            if e != 0:
                line = {'code':code, 'spec':spec, 'nr': e, 'total':total}
                if not provider in result:
                    result[provider] = [line]
                else:
                    result[provider].append(line)
        self._close()
        return result


if __name__ == "__main__":
    xp = XlsProcessor('./test.xlsx')
    orders = xp.gen_orders()
    import  utils
    print(utils.gen_all_orders_text(orders))

    up = [u'53d123', u'孙劲飞']
    print(xp.annotate_unknown_providers(up))
