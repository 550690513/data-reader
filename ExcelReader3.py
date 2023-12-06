import pandas as pd
import time
import os
import json
from openpyxl import load_workbook

# 读取Excel文件
workbook = load_workbook(filename="/Users/John/Documents/task_schedule.xlsx")
sheet = workbook["Sheet1"]

# 将Excel数据转换为DataFrame
df = pd.DataFrame(sheet.values)
headers = df.iloc[0]
df = pd.DataFrame(df.values[1:], columns=headers)

# 生成批量更新SQL语句
result = []
for index, row in df.iterrows():
    # print("正在处理第 %s 条数据" % (index + 1))
    content = row["content"]

    if content.find("AdminController#batchUpdate()") == -1:
        continue

    # print(content)

    # 使用 split() 方法将字符串分割成两部分
    split_parts = content.split("param=")
    if len(split_parts) > 1:
        param_part = split_parts[1]
    else:
        param_part = ""

    # print(param_part)

    # 解析 JSON 字符串
    data = json.loads(param_part)

    # 提取 ids 内容
    ids = data['request']['ids']

    # print(ids)

    for id in ids:
        result.append(id)

print(result)


