
import pandas as pd

file_path = '店端日报.xlsx'

# 先看看完整的Excel结构
df = pd.read_excel(file_path, header=None)

print('=' * 80)
print('Excel前20行, 前20列:')
print('=' * 80)
for i in range(min(20, len(df))):
    row = df.iloc[i].values[:20]
    print(f'行{i}:', [str(x)[:15] if pd.notna(x) else 'NaN' for x in row])

# 现在看看第2、3行作为列名
print('\n' + '=' * 80)
print('第1行(索引1):')
print('=' * 80)
print(df.iloc[1].values)

print('\n' + '=' * 80)
print('第2行(索引2):')
print('=' * 80)
print(df.iloc[2].values)

# 直接读取数据，第1、2行作为列名
print('\n' + '=' * 80)
print('尝试读取数据, 跳过前3行:')
print('=' * 80)
df_data = pd.read_excel(file_path, header=[1, 2], skiprows=0)
print(df_data.columns.tolist()[:30])

# 直接读取数据
df_data = pd.read_excel(file_path, header=[1, 2])
df_data = df_data.iloc[1:].copy()
print(f'\n数据形状:', df_data.shape)
print(df_data.head(3))

