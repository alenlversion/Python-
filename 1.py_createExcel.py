import openpyxl
from openpyxl.workbook import Workbook

# 1.创建excel文件
wb = openpyxl.Workbook()
# 2.默认sheet页
ws = wb.active
# 3.sheet页名称
ws.title = "成绩表"
# 4.创建表头
headers = ["姓名", "数学", "语文", "英语", "总分"]
# 添加表头到excel表中
ws.append(headers)

# 创建 数据
data = [
    ["小明", 12, 23, 54, ],
    ["王五", 33, 43, 54, ],
    ["赵六", 64, 86, 34, ],
    ["陈七", 86, 75, 87, ],
    ["黑旧", 85, 67, 86, ],
    ["是点开", 84, 96, 67, ],
]
# 添加数据到excel表中
for row in data:
    ws.append(row)

# 保存到excel表
wb.save("data/data.xlsx")
print("data save")
