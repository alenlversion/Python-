import numpy as np
import pandas as pd

data = {
    '科室': ['内科','外科','内科','外科','儿科','儿科'],
    '医生': ['A','B','C','A','D','E'],
    '收入': [1000, 2000, 1500, 1800, 800, 900],
    '成本': [400, 800, 600, 700, 300, 400],
    '日期': pd.to_datetime([
        '2025-01-01','2025-01-01','2025-01-02',
        '2025-01-02','2025-01-01','2025-01-02'
    ])
}

df = pd.DataFrame(data)

df['利润'] = df['收入'] - df['成本']
df_sr = df[df['收入']>1000]
df_count = df.groupby('科室')['科室'].value_counts()
# print(df_count)

df_agg = df.groupby('科室').agg(
    总收入=('收入','sum'),
    总成本=('成本','sum'),
    总利润=('利润','sum'),
)
#print(df_agg)

# df_max = df.loc[df.groupby('科室')['收入'].idxmax()]
# print(df_max)

# df_data_sum = df.groupby('日期')['收入'].sum().reset_index()
# print(df_data_sum)

# 制作透视表
df_table =  df.pivot_table(
    index='科室',
    columns='日期',
    values='收入',
    aggfunc='sum'
)

print(df_table)

# 制作 每个科室利润率
# df_profit = df.groupby('科室').agg(
#     总收入=('收入', 'sum'),
#     总成本=('成本', 'sum'),
#     总利润=('利润', 'sum'),
# )
# df_profit['利润率'] = df_profit['总利润'] / df_profit['总收入']
# print(df_profit.round(2))

# df_profit = df.groupby('科室').agg(
#     总收入=('收入', 'sum'),
#     总成本=('成本', 'sum'),
#     总利润=('利润', 'sum'),
# ).assign(
#     利润率=lambda x:x['总利润'] / x['总收入']
# )
#
# print(df_profit.round(2))

df['利润等级'] = df['利润'].apply(lambda x: '高收益' if x > 1000 else '普通')

df['收益等级'] = np.where(df['利润'] > 1000,'高收益','普通')
print(df)