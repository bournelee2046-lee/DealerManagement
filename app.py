
from flask import Flask, render_template, request, jsonify
import pandas as pd
import sqlite3
import os
from datetime import datetime, timedelta

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = './uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def get_db_connection():
    conn = sqlite3.connect('日报.db')
    conn.row_factory = sqlite3.Row
    return conn

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

def get_date_range(start_date_str, end_date_str):
    try:
        # 支持多种日期格式输入
        formats = ['%Y-%m-%d', '%Y/%m/%d']
        start_date = None
        end_date = None
        
        for fmt in formats:
            try:
                if not start_date:
                    start_date = datetime.strptime(start_date_str, fmt)
            except:
                pass
            try:
                if not end_date:
                    end_date = datetime.strptime(end_date_str, fmt)
            except:
                pass
        
        if not start_date or not end_date:
            return None
        
        dates = []
        current_date = start_date
        while current_date <= end_date:
            # 输出和数据库一致的格式：2026/4/30
            dates.append(f"{current_date.year}/{current_date.month}/{current_date.day}")
            current_date += timedelta(days=1)
        
        return dates
    except Exception as e:
        return None

def insert_data_to_db(df, dates, source, is_single_day):
    conn = get_db_connection()
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

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/daily_import')
def daily_import():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT 日报数据日期 FROM 日报快照 ORDER BY 日报数据日期 DESC")
    dates = cursor.fetchall()
    conn.close()
    
    return render_template('daily_import.html', existing_dates=[d['日报数据日期'] for d in dates])

@app.route('/follow_up')
def follow_up():
    return render_template('follow_up.html')

@app.route('/dates')
def get_dates():
    """获取所有可用日期，用于更新日期筛选器"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT 日报数据日期 FROM 日报快照 ORDER BY 日报数据日期 DESC")
    dates = cursor.fetchall()
    conn.close()
    
    return jsonify({'dates': [d['日报数据日期'] for d in dates]})

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '未选择文件'})
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'success': False, 'message': '文件名不能为空'})
    
    start_date = request.form.get('start_date')
    end_date = request.form.get('end_date')
    source = request.form.get('source', '零售部')
    
    if not start_date or not end_date:
        return jsonify({'success': False, 'message': '请选择日期范围'})
    
    dates = get_date_range(start_date, end_date)
    if not dates:
        return jsonify({'success': False, 'message': '日期格式无效'})
    
    # 判断是单日还是日期范围
    is_single_day = len(dates) == 1
    
    if file and allowed_file(file.filename):
        try:
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)
            
            df, error = parse_excel(file_path)
            
            if error:
                return jsonify({'success': False, 'message': error})
            
            inserted_count = insert_data_to_db(df, dates, source, is_single_day)
            
            os.remove(file_path)
            
            return jsonify({
                'success': True,
                'message': f'数据同步成功！共写入 {inserted_count} 条记录',
                'dates': dates
            })
        except Exception as e:
            return jsonify({'success': False, 'message': f'处理出错: {str(e)}'})
    
    return jsonify({'success': False, 'message': '不支持的文件格式'})

@app.route('/data')
def get_data():
    date_filter = request.args.get('date')
    search_text = request.args.get('search')
    limit = request.args.get('limit', 2000, type=int)  # 默认限制2000条
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 构建SQL查询
    sql = "SELECT * FROM 日报快照 WHERE 1=1"
    params = []
    
    if date_filter:
        sql += " AND 日报数据日期 = ?"
        params.append(date_filter)
    
    if search_text:
        sql += " AND (店编号 LIKE ? OR 店简称 LIKE ?)"
        search_pattern = f"%{search_text}%"
        params.append(search_pattern)
        params.append(search_pattern)
    
    sql += " ORDER BY 日报数据日期 DESC, 店编号"
    
    # 添加限制
    sql += " LIMIT ?"
    params.append(limit)
    
    cursor.execute(sql, params)
    rows = cursor.fetchall()
    conn.close()
    
    data = []
    for row in rows:
        data.append(dict(row))
    
    return jsonify({'data': data})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5003)

