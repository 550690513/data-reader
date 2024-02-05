import os
import time

import pandas as pd
from openpyxl import load_workbook

# 读取Excel文件
workbook = load_workbook(filename="/Users/John/Documents/华南华中春节调整后数据.xlsx")
sheet = workbook["Sheet2"]
# 将Excel数据转换为DataFrame
df = pd.DataFrame(sheet.values)

headers = df.iloc[0]
df = pd.DataFrame(df.values[1:], columns=headers)

# 生成批量更新SQL语句
updates = []
for index, row in df.iterrows():
    row_number = index + 1
    # print(f"{row_number}. 正在处理第 %s 条数据" % row_number)
    store_code = row["门店编码"]
    template_code = row["模板编码"]
    available_after = row["调整后订货日"]
    available_before = row["调整前订货日"]
    arrival_after = row["调整后到货日"]
    arrival_before = row["调整前到货日"]

    update_str = "UPDATE store_procurement_schedule SET"
    is_first = True
    prefix = " " if is_first else ", "
    if available_before != available_after:
        if available_before is not None:
            column = "available_time_excluded='[\"%s\"]'"
            value = pd.to_datetime(available_before).strftime("%Y-%m-%d")
            update_str = f"{update_str}{prefix}{column % value}"
            is_first = False
            prefix = " " if is_first else ", "
        if available_after is not None:
            column = "available_time_included='[\"%s\"]'"
            value = pd.to_datetime(available_after).strftime("%Y-%m-%d")
            update_str = f"{update_str}{prefix}{column % value}"
            is_first = False
            prefix = " " if is_first else ", "
    if arrival_before != arrival_after:
        if arrival_before is not None:
            column = "arrival_time_excluded='[\"%s\"]'"
            value = pd.to_datetime(arrival_before).strftime("%Y-%m-%d")
            update_str = f"{update_str}{prefix}{column % value}"
            is_first = False
            prefix = " " if is_first else ", "
        if arrival_after is not None:
            column = "arrival_time_included='[\"%s\"]'"
            value = pd.to_datetime(arrival_after).strftime("%Y-%m-%d")
            update_str = f"{update_str}{prefix}{column % value}"
            is_first = False
            prefix = " " if is_first else ", "
    where = " WHERE tenant_id = 1 AND template_code='%s' AND k3store_code='%s';"
    template_store = (template_code, store_code)
    update_str = f"{update_str}{where % template_store}"

    print(f"{row_number}. {update_str}")
    updates.append(update_str)

# 将SQL语句写入文件
timestamp = int(time.time())

# 将 SQL 语句写入文件
filename = f"result_{timestamp}.sql"
with open(filename, "w") as f:
    for update_str in updates:
        f.write(update_str + "\n")

# 打开 SQL 文件 (optional)
os.system(f"open {filename}")
