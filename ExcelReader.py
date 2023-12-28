import pandas as pd
import time
import os
from openpyxl import load_workbook

# 读取Excel文件
workbook = load_workbook(filename="/Users/John/Documents/元旦配送计划调整2023年12月25日.xlsx")
sheet = workbook["Sheet3"]

# 将Excel数据转换为DataFrame
df = pd.DataFrame(sheet.values)
headers = df.iloc[0]
df = pd.DataFrame(df.values[1:], columns=headers)

# 生成批量更新SQL语句
updates = []
for index, row in df.iterrows():
    print("正在处理第 %s 条数据" % (index + 1))
    store_code = row["门店编码"]
    template_code = row["模板编码"]
    available_after = row["调整后订货日"]
    available_before = row["调整前订货日"]
    arrival_after = row["调整后到货日"]
    arrival_before = row["调整前到货日"]

    # if "," in available_after:
    #     temp = [date for date in available_after.split(',')]
    # else:
    #     temp = [available_after]


    available_time_included = pd.to_datetime(available_after).strftime("%Y-%m-%d")
    available_time_excluded = pd.to_datetime(available_before).strftime("%Y-%m-%d")
    arrival_time_included = pd.to_datetime(arrival_after).strftime("%Y-%m-%d")
    arrival_time_excluded = pd.to_datetime(arrival_before).strftime("%Y-%m-%d")
    # available_time_included = available_after
    # available_time_excluded = available_before
    # arrival_time_included = arrival_after
    # arrival_time_excluded = arrival_before

    update = "UPDATE store_procurement_schedule SET available_time_included='[\"%s\"]', available_time_excluded='[\"%s\"]', arrival_time_included='[\"%s\"]', arrival_time_excluded='[\"%s\"]' WHERE tenant_id = 1 AND template_code='%s' AND k3store_code='%s';"
    values = (
        available_time_included,
        available_time_excluded,
        arrival_time_included,
        arrival_time_excluded,
        template_code,
        store_code,
    )
    update_str = update % values

    updates.append(update_str)

# 将SQL语句写入文件
timestamp = int(time.time())

# 将 SQL 语句写入文件
filename = f"result_{timestamp}.sql"
with open(filename, "w") as f:
    for update in updates:
        f.write(update + "\n")

# 打开 SQL 文件 (optional)
os.system(f"open {filename}")
