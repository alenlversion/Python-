import pandas as pd

df = pd.read_excel('C:/Users/Chen/Desktop/医生科室信息汇总表.xlsx')
# 1.筛选科室类型为临床科室
# result = df[df['科室类型'] == '临床科室']
# .query 方法 查询 科室类型为临床科室
result = df.query("科室类型 == '临床科室'")

# 2.查询出重复医生工号
duplicates_index = result[result.duplicated(subset=['医生名称'],keep=False)].index
# 3.对重复医生 进行排序展示
# duplicated_sorted = duplicates.sort_values(by=['医生名称','科室名称'])

def fow_styles(row):
    if row.name in duplicates_index:
        return ['background-color: yellow'] * len(row.index)
    else:
        return [''] * len(row)

styled_df = df.style.apply(fow_styles, axis=1)

styled_df.to_excel('C:/Users/Chen/Desktop/医生科室信息汇总表_已标记.xlsx',index=False,engine='openpyxl')


