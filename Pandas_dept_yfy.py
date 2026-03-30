import pandas as pd

# 1. 读取并筛选临床科室（关键：保留原始索引用于后续高亮）
df = pd.read_excel('C:/Users/Chen/Desktop/科室医生.xlsx')
# df_clinical = df[df['科室类型'] == '临床科室'].copy()  # 避免SettingWithCopyWarning

# 2. 【关键修正】确认重复判定列（根据您的注释"重复医生工号"）
# 请根据实际列名二选一：
#duplicate_col = '医生工号'      # ✅ 推荐：唯一标识，避免同名误判
duplicate_col = '原医生名称'  # ⚠️ 备用：仅当无工号字段且业务允许同名合并时使用

# 验证列是否存在
if duplicate_col not in df.columns:
    raise ValueError(f"错误：数据中不存在列 '{duplicate_col}'，请检查列名！可用列：{list(df.columns)}")

# 3. 标记重复行（基于筛选后的临床科室数据）
duplicates_mask = df.duplicated(subset=[duplicate_col], keep=False)
duplicates_index_set = set(df[duplicates_mask].index)  # 转为set提升查找效率

# 4. 按重复列+科室排序（使重复项在Excel中相邻显示）
df_sorted = df.sort_values(
    by=[duplicate_col],
    na_position='last',
    ignore_index=False  # 保留原始索引（关键！用于后续高亮匹配）
)

# 5. 高亮样式函数（使用set判断，效率更高）
def highlight_duplicates(row):
    return (
        ['background-color: yellow'] * len(row)
        if row.name in duplicates_index_set
        else [''] * len(row)
    )

# 6. 应用样式并导出
styled_df = df_sorted.style.apply(highlight_duplicates, axis=1)
output_path = 'C:/Users/Chen/Desktop/医生科室信息汇总表_已标记.xlsx'
styled_df.to_excel(output_path, index=False, engine='openpyxl')

# 7. 友好提示
print(f"✓ 处理完成！")
print(f"  • 标记重复项（按'{duplicate_col}'）: {len(duplicates_index_set)} 条")
print(f"  • 文件已保存: {output_path}")