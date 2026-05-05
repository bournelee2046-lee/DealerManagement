
import pandas as pd
import sqlite3
import os
from datetime import datetime, timedelta

def match_columns(excel_columns, db_columns):
    mapping = {}
    
    # 精确匹配数据库中已有的列名
    exact_matches = ['大区', '大区督导', '大区经理', '战区', '战区经理', '巡回员', '店编号', '店简称']
    for col in exact_matches:
        if col in excel_columns:
            mapping[col] = col
    
    # 模糊匹配其他字段
    field_keywords = {
        '30分钟跟进任务数': ['30分钟', '跟进任务'],
        '30分钟及时跟进数': ['30分钟及时', '及时跟进'],
        '三天三次跟进任务数': ['三天三次', '三天三次任务'],
        '三天三次跟进数': ['三天三次跟进'],
        '线索量-本地': ['线索量', '本地线索'],
        '到店数': ['到店数'],
        '到店数来源': ['到店数来源'],
        '数值性质': ['数值性质']
    }
    
    for db_col, keywords in field_keywords.items():
        if db_col in mapping.values():
            continue
        for excel_col in excel_columns:
            if excel_col in mapping:
                continue
            excel_col_str = str(excel_col)
            for keyword in keywords:
                if keyword in excel_col_str:
                    mapping[excel_col] = db_col
                    break
            if excel_col in mapping:
                break
    
    return mapping

def parse_excel(file_path):
    df = pd.read_excel(file_path, header=None)
    
    print('Excel前15行:')
    for i in range(min(15, len(df))):
        print(f'  行{i}: {df.iloc[i].values[:10]}')
    
    # 查找实际的数据起始行 - 大区列有数据的第一行
    data_start_row = -1
    for i in range(min(15, len(df))):
        cell_value = df.iloc[i, 0]
        if pd.notna(cell_value) and str(cell_value).strip() != '':
            # 检查是否是数据行（大区名称）
            val = str(cell_value).strip()
            if val not in ['大区', ''] and not pd.isna(val):
                data_start_row = i
                break
    
    if data_start_row == -1:
        print('未找到数据起始行')
        return None, '未找到数据起始行'
    
    print(f'\n找到数据起始在第 {data_start_row} 行')
    
    # 读取完整的数据，使用前两行作为多级表头，然后合并列名
    df_data = pd.read_excel(file_path, header=[0, 1])
    
    # 合并多级表头
    new_columns = []
    for col in df_data.columns:
        if pd.notna(col[1]) and str(col[1]).strip() != '':
            new_col = str(col[1]).strip()
        elif pd.notna(col[0]) and str(col[0]).strip() != '':
            new_col = str(col[0]).strip()
        else:
            new_col = f'Unnamed_{len(new_columns)}'
        new_columns.append(new_col)
    
    df_data.columns = new_columns
    
    print(f'\n合并后的列名: {new_columns[:30]}')
    
    # 跳过开头的空行和标题行，从实际数据行开始
    df_data = df_data.iloc[data_start_row - 2:].copy()
    
    # 删除空行
    df_data = df_data.dropna(how='all')
    
    # 重置索引
    df_data = df_data.reset_index(drop=True)
    
    print(f'\n数据行数: {len(df_data)}')
    print(f'\n前3条数据:')
    print(df_data.head(3))
    
    return df_data, None

def test_import():
    file_path = '店端日报.xlsx'
    
    if not os.path.exists(file_path):
        print(f'文件不存在: {file_path}')
        return
    
    print('=' * 50)
    print('开始测试导入...')
    print('=' * 50)
    
    df, error = parse_excel(file_path)
    
    if error:
        print(f'错误: {error}')
        return
    
    print('\n' + '=' * 50)
    print('查看数据库表结构')
    print('=' * 50)
    
    conn = sqlite3.connect('日报.db')
    cursor = conn.cursor()
    
    cursor.execute("PRAGMA table_info(日报快照)")
    db_columns_info = cursor.fetchall()
    db_columns = [col[1] for col in db_columns_info]
    print(f'数据库列: {db_columns}')
    
    excel_columns = df.columns.tolist()
    column_mapping = match_columns(excel_columns, db_columns)
    
    print(f'\n列映射关系:')
    for excel_col, db_col in column_mapping.items():
        print(f'  {excel_col} -> {db_col}')
    
    conn.close()
    
    print('\n' + '=' * 50)
    print('测试成功！文件可以正常解析')
    print('=' * 50)

if __name__ == '__main__':
    test_import()

