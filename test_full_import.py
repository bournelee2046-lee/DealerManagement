
import pandas as pd
import sqlite3
import os
import shutil
from datetime import datetime, timedelta

def match_columns(excel_columns, db_columns):
    mapping = {}
    
    # 精确匹配数据库中已有的列名
    exact_matches = ['大区', '大区督导', '大区经理', '战区', '战区经理', '巡回员', '店编号', '店简称']
    for col in exact_matches:
        if col in excel_columns:
            mapping[col] = col
    
    # 基于Excel中实际列名的匹配
    field_mapping = {
        '30分钟跟进任务数\n(首触)': '30分钟跟进任务数',
        '30分钟及时跟进数\n(首触)': '30分钟及时跟进数',
        '三天三次跟进任务数': '三天三次跟进任务数',
        '三天三次跟进数': '三天三次跟进数',
        '线索量\n（本地）': '线索量-本地',
        '到店数': '到店数'
    }
    
    for excel_col, db_col in field_mapping.items():
        if excel_col in excel_columns and db_col not in mapping.values():
            mapping[excel_col] = db_col
    
    return mapping

def parse_excel(file_path):
    # 使用前两行作为多级表头
    df_data = pd.read_excel(file_path, header=[1, 2])
    
    # 只保留"本月"下面的列和基本信息列
    columns_to_keep = []
    
    for col in df_data.columns:
        level1 = col[0] if pd.notna(col[0]) else ''
        level2 = col[1] if pd.notna(col[1]) else ''
        
        # 保留基本信息列（没有"本月"或"本日"的列）
        if level1 not in ['本月', '本日']:
            columns_to_keep.append(col)
        # 只保留"本月"下面的列
        elif level1 == '本月':
            columns_to_keep.append(col)
    
    df_data = df_data[columns_to_keep].copy()
    
    # 合并多级表头
    new_columns = []
    for col in df_data.columns:
        level1 = col[0] if pd.notna(col[0]) else ''
        level2 = col[1] if pd.notna(col[1]) else ''
        
        # 如果第二级有值，使用第二级
        if level2 and level2.strip() != '' and 'Unnamed' not in level2:
            new_columns.append(level2.strip())
        else:
            # 否则使用第一级
            new_columns.append(level1.strip())
    
    df_data.columns = new_columns
    
    # 不跳过任何行，所有行都是有效数据！
    # 只删除真正的全空行
    df_data = df_data.dropna(how='all')
    
    # 重置索引
    df_data = df_data.reset_index(drop=True)
    
    return df_data, None

def insert_data_to_db_test(df, dates, source, is_single_day):
    conn = sqlite3.connect('日报.db')
    cursor = conn.cursor()
    
    cursor.execute("PRAGMA table_info(日报快照)")
    db_columns_info = cursor.fetchall()
    db_columns = [col[1] for col in db_columns_info]
    
    excel_columns = df.columns.tolist()
    column_mapping = match_columns(excel_columns, db_columns)
    
    # 根据日期范围设置数值性质
    value_type = '连续值' if is_single_day else '区间值'
    
    inserted_count = 0
    for date_str in dates:
        for _, row in df.iterrows():
            values = {}
            for excel_col, db_col in column_mapping.items():
                value = row.get(excel_col)
                if pd.isna(value):
                    values[db_col] = None
                else:
                    values[db_col] = str(value)
            
            # 设置日报数据日期
            values['日报数据日期'] = date_str
            
            # 设置到店数来源和数值性质
            values['到店数来源'] = source
            values['数值性质'] = value_type
            
            columns = list(values.keys())
            placeholders = ', '.join(['?' for _ in columns])
            quoted_columns = ', '.join([f'"{col}"' for col in columns])
            sql = f'INSERT INTO 日报快照 ({quoted_columns}) VALUES ({placeholders})'
            
            cursor.execute(sql, list(values.values()))
            inserted_count += 1
    
    conn.commit()
    conn.close()
    
    return inserted_count

def test_full_import():
    file_path = '店端日报.xlsx'
    
    print('=' * 80)
    print('第一步：测试单日导入（数值性质应为连续值）')
    print('=' * 80)
    
    df_data, error = parse_excel(file_path)
    if error:
        print(error)
        return
    
    print(f'解析成功，数据形状：{df_data.shape}')
    
    # 备份数据库
    if os.path.exists('日报.db.bak'):
        os.remove('日报.db.bak')
    shutil.copy('日报.db', '日报.db.bak')
    
    try:
        # 测试单日导入
        conn = sqlite3.connect('日报.db')
        cursor = conn.cursor()
        
        # 删除测试日期的数据
        cursor.execute("DELETE FROM 日报快照 WHERE 日报数据日期 = '2026/5/1'")
        conn.commit()
        conn.close()
        
        # 单日导入
        inserted_count = insert_data_to_db_test(df_data, ['2026/5/1'], '打铁系统', True)
        print(f'成功插入 {inserted_count} 条数据')
        
        # 验证
        conn = sqlite3.connect('日报.db')
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM 日报快照 WHERE 日报数据日期 = '2026/5/1'")
        count = cursor.fetchone()[0]
        print(f'验证：日期2026/5/1共有 {count} 条记录')
        
        cursor.execute("SELECT * FROM 日报快照 WHERE 日报数据日期 = '2026/5/1' LIMIT 1")
        row = cursor.fetchone()
        print(f'第一条记录的到店数来源：{row[15]}，数值性质：{row[16]}')
        conn.close()
        
    except Exception as e:
        print(f'出错：{e}')
        import traceback
        traceback.print_exc()
        # 恢复数据库
        if os.path.exists('日报.db'):
            os.remove('日报.db')
        shutil.copy('日报.db.bak', '日报.db')
        os.remove('日报.db.bak')
        return
    
    print('\n' + '=' * 80)
    print('第二步：测试日期范围导入（数值性质应为区间值）')
    print('=' * 80)
    
    try:
        # 删除测试日期的数据
        conn = sqlite3.connect('日报.db')
        cursor = conn.cursor()
        cursor.execute("DELETE FROM 日报快照 WHERE 日报数据日期 IN ('2026/4/23', '2026/4/24', '2026/4/25')")
        conn.commit()
        conn.close()
        
        # 日期范围导入
        inserted_count = insert_data_to_db_test(df_data, ['2026/4/23', '2026/4/24', '2026/4/25'], '零售部', False)
        print(f'成功插入 {inserted_count} 条数据')
        
        # 验证
        conn = sqlite3.connect('日报.db')
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM 日报快照 WHERE 日报数据日期 IN ('2026/4/23', '2026/4/24', '2026/4/25')")
        count = cursor.fetchone()[0]
        print(f'验证：日期范围2026/4/23-2026/4/25共有 {count} 条记录')
        
        for date in ['2026/4/23', '2026/4/24', '2026/4/25']:
            cursor.execute(f"SELECT * FROM 日报快照 WHERE 日报数据日期 = '{date}' LIMIT 1")
            row = cursor.fetchone()
            print(f'日期{date}第一条记录：到店数来源={row[15]},数值性质={row[16]}')
        
        conn.close()
        
    except Exception as e:
        print(f'出错：{e}')
        import traceback
        traceback.print_exc()
        # 恢复数据库
        if os.path.exists('日报.db'):
            os.remove('日报.db')
        shutil.copy('日报.db.bak', '日报.db')
        os.remove('日报.db.bak')
        return
    
    # 恢复原始数据
    os.remove('日报.db')
    shutil.copy('日报.db.bak', '日报.db')
    os.remove('日报.db.bak')
    
    print('\n' + '=' * 80)
    print('所有测试完成！')
    print('=' * 80)

if __name__ == '__main__':
    test_full_import()

