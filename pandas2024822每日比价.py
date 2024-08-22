import pandas as pd
import os

# 文件夹路径
folder_path = f'C:\\Users\\Administrator\\Desktop\\pandas\\每日均价\\'

# 获取文件夹中所有以"综合成本核算表"开头的Excel文件的文件名列表
file_names = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.startswith('综合成本核算表') and f.endswith('.xlsx')]

# 初始化列表来存放所有的DataFrame
dfs = []
dates = []

# 读取所有Excel文件并存入列表，同时提取文件名中的日期信息
for file_name in file_names:
    df = pd.read_excel(file_name)
    dfs.append(df)
    # 提取文件名中的日期部分
    date = os.path.splitext(os.path.basename(file_name))[0].replace('综合成本核算表', '')
    dates.append(date)

# 提取并合并所有DataFrame的商品代码，删除重复值
all_product_codes = pd.concat([df['商品代码'] for df in dfs]).drop_duplicates().reset_index(drop=True)

# 计算每个商品代码的销售数量合计
sales_sums = {code: 0 for code in all_product_codes}

for df in dfs:
    for code in all_product_codes:
        if code in df['商品代码'].values:
            sales_sums[code] += df[df['商品代码'] == code]['本日销售带皮'].sum()

# 按销售数量合计从大到小排序
sorted_product_codes = sorted(sales_sums.items(), key=lambda x: x[1], reverse=True)
sorted_codes = [code for code, _ in sorted_product_codes]
sorted_sales_sums = [sum_ for _, sum_ in sorted_product_codes]

# 获取商品名称
product_names = {}
for df in dfs:
    for code in sorted_codes:
        if code in df['商品代码'].values:
            product_names[code] = df[df['商品代码'] == code]['商品名称'].values[0]

# 创建一个新的DataFrame来存放结果，第一行为销售数量合计，第二行为商品代码，第三行为商品名称
result_df = pd.DataFrame(columns=['日期'] + sorted_codes)

# 将销售数量合计添加到结果DataFrame的第一行
result_df.loc[0] = ['销售数量合计'] + sorted_sales_sums

# 将商品代码添加到第二行
result_df.loc[1] = ['商品代码'] + sorted_codes

# 将商品名称添加到第三行
result_df.loc[2] = ['商品名称'] + [product_names[code] for code in sorted_codes]

# 匹配每个商品在每个日期的单价并添加到result_df的对应行
cost_prices = {code: [] for code in sorted_codes}  # 用于存储每个商品代码的所有成本价
for i, (df, date) in enumerate(zip(dfs, dates)):
    prices = []
    for code in sorted_codes:
        price = df[df['商品代码'] == code]['本日核算成本价'].values
        if len(price) > 0:
            price = price[0]
            cost_prices[code].append(price)  # 添加到该商品代码的成本价列表
        else:
            price = None
        prices.append(price)
    result_df.loc[date] = [date] + prices

# 将倒数第二次成本价添加到最后一行
second_last_prices = ['倒数第二次成本价'] + [cost_prices[code][-2] if len(cost_prices[code]) >= 2 else None for code in sorted_codes]
result_df.loc['倒数第二次成本价'] = second_last_prices

# 打印结果DataFrame以查看结果
print(result_df)

# 保存结果到新的Excel文件
result_df.to_excel(os.path.join(folder_path, '综合成本核算结果.xlsx'), index=False)