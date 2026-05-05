
import pandas as pd
import sqlite3

# 查看Excel中的数据
print('=' * 80)
print('分析Excel原始数据')
print('=' * 80)

df_raw = pd.read_excel('店端日报.xlsx', header=None)
print(f'Excel总行数: {len(df_raw)}')
print(f'\n前4行:')
print(df_raw.iloc[:4])

print(f'\n从第3行开始的数据行数: {len(df_raw.iloc[3:])}')

# 用header=[1,2]读取，然后数一下
df = pd.read_excel('店端日报.xlsx', header=[1, 2])
print(f'\n用多级表头读取，总行数: {len(df)}')
print(f'跳过第1行后的行数: {len(df.iloc[1:])}')

print('\n' + '=' * 80)
print('检查第0行数据(被跳过的)')
print('=' * 80)
print(df.iloc[0, :9])

# 检查数据库中的数据
print('\n' + '=' * 80)
print('检查数据库中的数据')
print('=' * 80)
conn = sqlite3.connect('日报.db')
cursor = conn.cursor()
cursor.execute("SELECT COUNT(*) FROM 日报快照 WHERE 日报数据日期 = '2026/4/30'")
count = cursor.fetchone()[0]
print(f'数据库中2026/4/30的数据行数: {count}')

# 检查我们的代码逻辑
print('\n' + '=' * 80)
print('让我们用实际数据再试一次')
print('=' * 80)

df_data = pd.read_excel('店端日报.xlsx', header=[1, 2])

# 只保留"本月"下面的列和基本信息列
columns_to_keep = []
for col in df_data.columns:
    level1 = col[0] if pd.notna(col[0]) else ''
    level2 = col[1] if pd.notna(col[1]) else ''
    if level1 not in ['本月', '本日']:
        columns_to_keep.append(col)
    elif level1 == '本月':
        columns_to_keep.append(col)

df_data = df_data[columns_to_keep].copy()

# 合并多级表头
new_columns = []
for col in df_data.columns:
    level1 = col[0] if pd.notna(col[0]) else ''
    level2 = col[1] if pd.notna(col[1]) else ''
    if level2 and level2.strip() != '' and 'Unnamed' not in level2:
        new_columns.append(level2.strip())
    else:
        new_columns.append(level1.strip())
df_data.columns = new_columns

print(f'处理后的列名: {new_columns[:9]}')

# 看看原始数据
print(f'\n原始df_data长度: {len(df_data)}')
print(f'第0行:')
print(df_data.iloc[0, :9])
print(f'\n第1行:')
print(df_data.iloc[1, :9])

# 检查第0行是否为空
print(f'\n第0行是否全空: {df_data.iloc[0].isna().all()}')
print(f'第0行的大区列: {df_data.iloc[0, 0]}')

# 看看不跳过的话有多少行
df_full = df_data.copy()
print(f'\n不跳过任何行的数据形状: {df_full.shape}')
print(f'全空的行数: {df_full.isna().all(axis=1).sum()}')
print(f'全空的行索引: {df_full.index[df_full.isna().all(axis=1)].tolist()}')

# 看看不删除全空行的数据
df_no_drop = df_data.iloc[1:].copy()
print(f'\n只跳过第0行的数据形状: {df_no_drop.shape}')

print(f'\n第0行的数据:')
print(df_data.iloc[0, :9])

conn.close()

