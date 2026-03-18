import pandas as pd


# ==========================================
# 准备测试数据 (模拟图书馆借阅与库存)
# ==========================================
def get_test_data():
    # 图书详情表
    books_data = {
        '书名': ['Python基础', '数据结构', '数据库原理', 'Kettle实战', '机器学习'],
        '单价': [58.0, 45.0, 62.0, 39.0, 88.0],
        '库存': [10, 0, 5, 8, 3],
        '分类': ['编程', '基础', '数据库', '工具', '人工智能']
    }
    df_books = pd.DataFrame(books_data)

    # 借阅数量 (Series) - 注意索引与 df_books 的行标签对应
    # 模拟一部分书籍有借阅记录
    borrow_series = pd.Series([2, 5, 1], index=[0, 2, 4], name='借阅数')

    return df_books, borrow_series



