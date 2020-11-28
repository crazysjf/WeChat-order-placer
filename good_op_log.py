import pandas as pd

class GoodsOpLog():
    '''处理商品资料的类'''

    def __init__(self, f):
        self._df = pd.read_excel(f)

    def test(self):
        df = self._df
        df = df[df['操作类型']=="快速上架"]
        df = df.groupby("商品编码")['数量'].sum()
        print(df)
