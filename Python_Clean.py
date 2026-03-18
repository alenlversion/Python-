import pandas as pd

'''
DataFrame 包含删除空字段的行 dropna()方法
DataFame.dropna(axis=0,how='any',thresh=None,subset=None,inplace=False )
axis 默认值为 0 代表遇到空值就剔除， 为 1 表示 遇到空值 剔除整列 
how 默认值 any  如果行/列里任何一个数据出现NA就去掉整行,如果设置 how='all' 行和列 都是NA 才会去掉
thresh：设置需要多少非空值的数据才可以保留下来的。
subset：设置想要检查的列。如果是多个列，可以使用列名的 list 作为参数。
inplace：如果设置 True，将计算得到的值直接覆盖之前的值并返回 None，修改的是源数据。
'''

# df = pd.read_csv('data/property-data.csv')
# print(df['NUM_BEDROOMS'])

# missing_values = ["n/a","na","--"]
# df = pd.read_csv('data/property-data.csv', na_values = missing_values)
#
# print(df['NUM_BEDROOMS'])
# print(df['NUM_BEDROOMS'].isnull())

df = pd.read_csv('data/property-data.csv')

# dropDf = df.dropna()
#
# print(df.to_string())
# print(dropDf.to_string())

# # 移除指定列的空行  ST_NUM
# drop_St_Num = df.dropna(subset=['ST_NUM'],inplace=False)
#
# print(drop_St_Num)

# # 将空值替换为 其他内容
# fillna_St_Num = df.fillna(12345,inplace=False)
#
# print(fillna_St_Num)

# # 将某个字段的空值 进行替换
# fillna_PID = df['PID'].fillna('12345',inplace=False)
# print(fillna_PID)

ST_NUM_Mean = df["ST_NUM"].mean()
print(df)
print(ST_NUM_Mean)

