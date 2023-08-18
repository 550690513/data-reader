import pandas as pd
import time
import os
from openpyxl import load_workbook

# 读取Excel文件
workbook = load_workbook(filename="/Users/John/Documents/data_qixi.xlsx")
sheet = workbook["Sheet1"]

# 将Excel数据转换为DataFrame
df = pd.DataFrame(sheet.values)
headers = df.iloc[0]
df = pd.DataFrame(df.values[1:], columns=headers)

# 生成批量更新SQL语句
sql = []
condition = []
for index, row in df.iterrows():
    template_code = row["模板编码"]
    store_code = row["门店编码"]
    template_stores = {}
    template_stores[template_code] = store_code
    condition.append(template_stores)

for index, item in enumerate(condition):
    select = f"SELECT k3store_code, template_code, arrival_days, available_time_included, available_time_excluded, arrival_time_included, arrival_time_excluded " \
             f"FROM store_procurement_schedule " \
             f"WHERE tenant_id = 1 AND template_code = '{list(item.keys())[0]}' AND k3store_code = '{list(item.values())[0]}'"
    if index != 0:
        sql.append("UNION ALL")
    sql.append(select)
    if index == len(condition) - 1:
        sql.append(";")

# 将SQL语句写入文件
timestamp = int(time.time())

# 将 SQL 语句写入文件
filename = f"result_{timestamp}.sql"
with open(filename, "w") as f:
    for select in sql:
        f.write(select + "\n")

# 打开 SQL 文件 (optional)
os.system(f"open {filename}")
