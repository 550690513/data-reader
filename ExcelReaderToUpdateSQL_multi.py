import pandas as pd
import time
import os
from openpyxl import load_workbook

# 读取Excel文件
workbook = load_workbook(filename="/Users/John/Documents/高佳倩数据.xlsx")
sheet = workbook["Sheet4"]
# 将Excel数据转换为DataFrame
df = pd.DataFrame(sheet.values)

headers = df.iloc[0]
df = pd.DataFrame(df.values[1:], columns=headers)

# 生成批量更新SQL语句
updates = []
for index, row in df.iterrows():
    row_number = (index + 1)
    print(f"{row_number}. 正在处理第 %s 条数据" % row_number)

    template_code = row["模板编码"]
    store_code = row["门店编码"]
    available_before = row["调整前订货日"]
    available_after = row["调整后订货日"]
    arrival_before = row["调整前到货日"]
    arrival_after = row["调整后到货日"]

    # 多日期的，单独处理
    if "," in available_before:
        available_before = [pd.to_datetime(date).strftime("%Y-%m-%d") for date in available_before.split(',')]
    else:
        available_before = [pd.to_datetime(available_before).strftime("%Y-%m-%d")]
    if "," in available_after:
        available_after = [pd.to_datetime(date).strftime("%Y-%m-%d") for date in available_after.split(',')]
    else:
        available_after = [pd.to_datetime(available_after).strftime("%Y-%m-%d")]
    if "," in arrival_before:
        arrival_before = [pd.to_datetime(date).strftime("%Y-%m-%d") for date in arrival_before.split(',')]
    else:
        arrival_before = [pd.to_datetime(arrival_before).strftime("%Y-%m-%d")]
    if "," in arrival_after:
        arrival_after = [pd.to_datetime(date).strftime("%Y-%m-%d") for date in arrival_after.split(',')]
    else:
        arrival_after = [pd.to_datetime(arrival_after).strftime("%Y-%m-%d")]

    available_time_excluded = '","'.join(available_before)
    available_time_included = '","'.join(available_after)
    arrival_time_excluded = '","'.join(arrival_before)
    arrival_time_included = '","'.join(arrival_after)

    update = ""
    values = ()
    if available_time_excluded != available_time_included and arrival_time_excluded != arrival_time_included:
        print(f"{row_number}. 订货日、到货日均调整...")
        update = "UPDATE store_procurement_schedule SET available_time_excluded='[\"%s\"]', available_time_included='[\"%s\"]', arrival_time_excluded='[\"%s\"]', arrival_time_included='[\"%s\"]' WHERE tenant_id = 1 AND template_code='%s' AND k3store_code='%s';"
        values = (
            available_time_excluded,
            available_time_included,
            arrival_time_excluded,
            arrival_time_included,
            template_code,
            store_code,
        )
    elif available_time_excluded == available_time_included and arrival_time_excluded != arrival_time_included:
        print(f"{row_number}. 订货日不调整，到货日调整...")
        update = "UPDATE store_procurement_schedule SET arrival_time_excluded='[\"%s\"]', arrival_time_included='[\"%s\"]' WHERE tenant_id = 1 AND template_code='%s' AND k3store_code='%s';"
        values = (
            arrival_time_excluded,
            arrival_time_included,
            template_code,
            store_code,
        )
    elif available_time_excluded != available_time_included and arrival_time_excluded == arrival_time_included:
        print(f"{row_number}. 订货日调整，到货日不调整...")
        update = "UPDATE store_procurement_schedule SET available_time_excluded='[\"%s\"]', available_time_included='[\"%s\"]' WHERE tenant_id = 1 AND template_code='%s' AND k3store_code='%s';"
        values = (
            available_time_excluded,
            available_time_included,
            template_code,
            store_code,
        )
    else:
        print(f"{row_number}. 订货日、到货日均不调整，ignore...")

    update_str = update % values
    print(f"{row_number}. {update_str}")

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
