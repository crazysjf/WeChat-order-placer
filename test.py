import pandas as pd

import math
import os
import pandas
raw_data = {
        '供应商': ['1', '2', '1', '2', '5'],
        '供应商商品款号': ['1186', '567', '1186', '567', 'Ayoung'],
        '数量': ['1', '2', '3', '4', '5']}
df = pd.DataFrame(raw_data, columns = ['供应商', '供应商商品款号', '数量'])

#print(df)

df2 = pd.DataFrame(columns = ['供应商', '供应商商品款号', '数量'])

#print(df[df['供应商']=='1'])

df3 = df.drop_duplicates(subset=['供应商','供应商商品款号'],keep='first')
#print(df3)

print("----")
for r in df3.index:
    provider = df3.loc[r]["供应商"]
    code = df3.loc[r]["供应商商品款号"]
    df_tmp = df[(df['供应商']==provider) & (df['供应商商品款号']==code)]
    print(df_tmp)
    sum = df_tmp['数量'].apply(lambda x: int(x)).sum()
    df3.loc[r]['数量'] = sum

print(df3)