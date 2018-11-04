import pandas as pd

class GoodsProfile():
    '''处理商品资料的类'''

    def __init__(self, f):
        self._df = pd.read_excel(f)

    def get_df(self):
        return self._df
    # def _get_xls_column_nr(self, table_head, name):
    #     for i in range(0, len(table_head)):
    #         if table_head[i] == name:
    #             return i
    #     return None