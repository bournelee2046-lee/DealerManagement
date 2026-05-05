
from flask import Flask, render_template, request, jsonify
import pandas as pd
import sqlite3
import os
import json
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
@app.route('/api/dates')
def get_dates():
    """获取所有可用日期，用于更新日期筛选器"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT 日报数据日期 FROM 日报快照 ORDER BY 日报数据日期 DESC")
    dates = cursor.fetchall()
    conn.close()
    
    return jsonify({'success': True, 'data': [{'date': d['日报数据日期']} for d in dates]})


@app.route('/api/stores')
def get_stores():
    """获取门店统计数据（兼容旧API路径）"""
    date_filter = request.args.get('date')
    
    if not date_filter:
        return jsonify({'success': False, 'message': '请选择日期'})
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT 店编号, 店简称, 大区, 战区,
               COALESCE(MAX(CAST("线索量-本地" AS REAL)), 0) as 线索量本地,
               COALESCE(MAX(CAST(到店数 AS REAL)), 0) as 到店数
        FROM 日报快照
        WHERE 日报数据日期 = ?
        GROUP BY 店编号, 店简称, 大区, 战区
        ORDER BY 店编号
    """, (date_filter,))
    
    rows = cursor.fetchall()
    conn.close()
    
    data = []
    seen = set()
    for row in rows:
        record = dict(row)
        store_code = record['店编号']
        if store_code and store_code not in seen:
            seen.add(store_code)
            try:
                线索量 = float(record['线索量本地'])
                到店数 = float(record['到店数'])
                if 线索量 > 0:
                    record['到店率'] = round(到店数 / 线索量 * 100, 2)
                else:
                    record['到店率'] = 0
            except:
                record['到店率'] = 0
            data.append(record)
    
    return jsonify({'success': True, 'data': data})

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
    limit = request.args.get('limit', 2000, type=int)
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
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
    
    sql += " LIMIT ?"
    params.append(limit)
    
    cursor.execute(sql, params)
    rows = cursor.fetchall()
    conn.close()
    
    data = []
    for row in rows:
        data.append(dict(row))
    
    return jsonify({'data': data})

@app.route('/api/store_stats')
def get_store_stats():
    """获取门店统计数据（用于门店选择页面）"""
    date_filter = request.args.get('date')
    
    if not date_filter:
        return jsonify({'success': False, 'message': '请选择日期'})
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT 店编号, 店简称, 大区, 战区,
               COALESCE("线索量-本地", 0) as 线索量本地,
               COALESCE(到店数, 0) as 到店数
        FROM 日报快照
        WHERE 日报数据日期 = ?
        ORDER BY 店编号
    """, (date_filter,))
    
    rows = cursor.fetchall()
    conn.close()
    
    data = []
    for row in rows:
        record = dict(row)
        try:
            线索量 = float(record['线索量本地'])
            到店数 = float(record['到店数'])
            if 线索量 > 0:
                record['到店率'] = round(到店数 / 线索量 * 100, 2)
            else:
                record['到店率'] = 0
        except:
            record['到店率'] = 0
        data.append(record)
    
    return jsonify({'success': True, 'data': data})

@app.route('/api/store_trend')
def get_store_trend():
    """获取单个门店的数据变化趋势"""
    store_code = request.args.get('store_code')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    
    if not store_code:
        return jsonify({'success': False, 'message': '缺少店编号'})
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    sql = """
        SELECT 日报数据日期, "线索量-本地", 到店数
        FROM 日报快照
        WHERE 店编号 = ?
    """
    params = [store_code]
    
    if start_date:
        sql += " AND 日报数据日期 >= ?"
        params.append(start_date)
    
    if end_date:
        sql += " AND 日报数据日期 <= ?"
        params.append(end_date)
    
    sql += " ORDER BY 日报数据日期"
    
    cursor.execute(sql, params)
    rows = cursor.fetchall()
    conn.close()
    
    data = []
    for row in rows:
        record = dict(row)
        try:
            线索量 = float(record['线索量-本地']) if record['线索量-本地'] else 0
            到店数 = float(record['到店数']) if record['到店数'] else 0
            record['线索量数值'] = 线索量
            record['到店数数值'] = 到店数
            if 线索量 > 0:
                record['到店率'] = round(到店数 / 线索量 * 100, 2)
            else:
                record['到店率'] = 0
        except:
            record['线索量数值'] = 0
            record['到店数数值'] = 0
            record['到店率'] = 0
        data.append(record)
    
    return jsonify({'success': True, 'data': data})

@app.route('/api/follow_reasons', methods=['GET', 'POST', 'PUT', 'DELETE'])
@app.route('/api/follow_reason', methods=['GET', 'POST', 'PUT', 'DELETE'])
def manage_follow_reasons():
    """管理跟进原因配置（支持二级分类）"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if request.method == 'GET':
        cursor.execute("""
            SELECT 配置ID, 原因选项, 排序, 状态, 父级ID
            FROM 跟进原因配置
            ORDER BY 父级ID, 排序
        """)
        all_reasons = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        # 转换为前端期望的格式（扁平化列表
        result = []
        for reason in all_reasons:
            result.append({
                'id': reason['配置ID'],
                '原因': reason['原因选项'],
                '父级ID': reason['父级ID'],
                '原因选项': reason['原因选项'],
                '配置ID': reason['配置ID']
            })
        
        return jsonify({'success': True, 'data': result})
    
    elif request.method == 'POST':
        data = request.get_json()
        reason = data.get('原因选项', '').strip()
        parent_id = data.get('父级ID', 0)
        
        if not reason:
            conn.close()
            return jsonify({'success': False, 'message': '原因选项不能为空'})
        
        cursor.execute("SELECT MAX(排序) as max_sort FROM 跟进原因配置 WHERE 父级ID = ?", (parent_id,))
        max_sort = cursor.fetchone()['max_sort'] or 0
        
        cursor.execute("""
            INSERT INTO 跟进原因配置 (原因选项, 排序, 状态, 父级ID, 创建时间)
            VALUES (?, ?, '启用', ?, datetime('now'))
        """, (reason, max_sort + 1, parent_id))
        
        config_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': '添加成功', '配置ID': config_id})
    
    elif request.method == 'PUT':
        data = request.get_json()
        config_id = data.get('配置ID')
        reason = data.get('原因选项', '').strip()
        status = data.get('状态')
        sort_order = data.get('排序')
        parent_id = data.get('父级ID')
        
        if not config_id:
            conn.close()
            return jsonify({'success': False, 'message': '缺少配置ID'})
        
        updates = []
        params = []
        if reason:
            updates.append("原因选项 = ?")
            params.append(reason)
        if status:
            updates.append("状态 = ?")
            params.append(status)
        if sort_order is not None:
            updates.append("排序 = ?")
            params.append(sort_order)
        if parent_id is not None:
            updates.append("父级ID = ?")
            params.append(parent_id)
        
        if updates:
            params.append(config_id)
            cursor.execute(f"""
                UPDATE 跟进原因配置
                SET {', '.join(updates)}
                WHERE 配置ID = ?
            """, params)
            conn.commit()
        
        conn.close()
        return jsonify({'success': True, 'message': '更新成功'})
    
    elif request.method == 'DELETE':
        config_id = request.args.get('配置ID')
        if not config_id:
            conn.close()
            return jsonify({'success': False, 'message': '缺少配置ID'})
        
        cursor.execute("DELETE FROM 跟进原因配置 WHERE 父级ID = ?", (config_id,))
        cursor.execute("DELETE FROM 跟进原因配置 WHERE 配置ID = ?", (config_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': '删除成功'})

@app.route('/api/follow_tasks', methods=['GET', 'POST'])
def manage_follow_tasks():
    """管理跟进任务"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if request.method == 'GET':
        cursor.execute("""
            SELECT 任务ID, 周开始日期, 门店列表, 状态, 创建时间
            FROM 跟进任务
            ORDER BY 创建时间 DESC
        """)
        tasks = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        # 转换为前端期望的格式
        formatted_tasks = []
        for task in tasks:
            formatted_tasks.append({
                'id': task['任务ID'],
                'title': f"{task['周开始日期']} 跟进任务",
                'dateRange': task['周开始日期'],
                '周开始日期': task['周开始日期'],
                '门店列表': task['门店列表'],
                '状态': task['状态']
            })
        
        return jsonify({'success': True, 'data': formatted_tasks})
    
    elif request.method == 'POST':
        data = request.get_json()
        week_start = data.get('周开始日期') or data.get('title')
        store_list = data.get('门店列表', '[]')
        
        if not week_start:
            conn.close()
            return jsonify({'success': False, 'message': '缺少周开始日期'})
        
        cursor.execute("""
            INSERT INTO 跟进任务 (周开始日期, 门店列表, 状态, 创建时间)
            VALUES (?, ?, '进行中', datetime('now'))
        """, (week_start, store_list))
        
        task_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'data': {'id': task_id}, 'message': '任务创建成功'})

@app.route('/api/follow_tasks/<int:task_id>', methods=['GET', 'PUT', 'DELETE'])
def manage_follow_task(task_id):
    """单个跟进任务的操作"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if request.method == 'GET':
        cursor.execute("""
            SELECT 任务ID, 周开始日期, 门店列表, 状态, 创建时间
            FROM 跟进任务
            WHERE 任务ID = ?
        """, (task_id,))
        task = cursor.fetchone()
        
        if not task:
            conn.close()
            return jsonify({'success': False, 'message': '任务不存在'})
        
        task_data = dict(task)
        
        # 解析门店列表并获取门店详情
        store_list_str = task_data['门店列表']
        stores = []
        seen = set()
        
        if store_list_str:
            try:
                # 尝试解析JSON格式
                import json
                store_codes = json.loads(store_list_str)
                if isinstance(store_codes, list) and len(store_codes) > 0 and isinstance(store_codes[0], dict):
                    # 如果是对象数组，去重
                    for store in store_codes:
                        code = store.get('店编号')
                        if code and code not in seen:
                            seen.add(code)
                            stores.append(store)
                else:
                    # 如果是字符串列表，获取门店详情
                    if isinstance(store_codes, list):
                        codes = store_codes
                    else:
                        codes = str(store_list_str).split(',')
                    
                    if codes:
                        # 先去重门店编号
                        unique_codes = []
                        code_seen = set()
                        for code in codes:
                            if code and code not in code_seen:
                                code_seen.add(code)
                                unique_codes.append(code)
                        
                        if unique_codes:
                            placeholders = ','.join(['?' for _ in unique_codes])
                            cursor.execute(f"""
                                SELECT 店编号, 店简称, 大区, 战区,
                                       COALESCE(MAX(CAST("线索量-本地" AS REAL)), 0) as 线索量本地,
                                       COALESCE(MAX(CAST(到店数 AS REAL)), 0) as 到店数
                                FROM 日报快照
                                WHERE 店编号 IN ({placeholders})
                                GROUP BY 店编号, 店简称, 大区, 战区
                                ORDER BY 店编号
                            """, unique_codes)
                            
                            for row in cursor.fetchall():
                                store = dict(row)
                                try:
                                    线索量 = float(store['线索量本地'])
                                    到店数 = float(store['到店数'])
                                    if 线索量 > 0:
                                        store['到店率'] = round(到店数 / 线索量 * 100, 2)
                                    else:
                                        store['到店率'] = 0
                                except:
                                    store['到店率'] = 0
                                stores.append(store)
            except:
                # 如果不是JSON，尝试按逗号分割
                codes = str(store_list_str).split(',')
                if codes and codes[0]:
                    # 先去重门店编号
                    unique_codes = []
                    code_seen = set()
                    for code in codes:
                        if code and code not in code_seen:
                            code_seen.add(code)
                            unique_codes.append(code)
                    
                    if unique_codes:
                        placeholders = ','.join(['?' for _ in unique_codes])
                        cursor.execute(f"""
                            SELECT 店编号, 店简称, 大区, 战区,
                                   COALESCE(MAX(CAST("线索量-本地" AS REAL)), 0) as 线索量本地,
                                   COALESCE(MAX(CAST(到店数 AS REAL)), 0) as 到店数
                            FROM 日报快照
                            WHERE 店编号 IN ({placeholders})
                            GROUP BY 店编号, 店简称, 大区, 战区
                            ORDER BY 店编号
                        """, unique_codes)
                        
                        for row in cursor.fetchall():
                            store = dict(row)
                            try:
                                线索量 = float(store['线索量本地'])
                                到店数 = float(store['到店数'])
                                if 线索量 > 0:
                                    store['到店率'] = round(到店数 / 线索量 * 100, 2)
                                else:
                                    store['到店率'] = 0
                            except:
                                store['到店率'] = 0
                            stores.append(store)
        
        conn.close()
        
        return jsonify({
            'success': True,
            'data': {
                'id': task_data['任务ID'],
                'title': f"{task_data['周开始日期']} 跟进任务",
                'dateRange': task_data['周开始日期'],
                '周开始日期': task_data['周开始日期'],
                '门店列表': task_data['门店列表'],
                'stores': stores,
                '状态': task_data['状态']
            }
        })
    
    elif request.method == 'PUT':
        data = request.get_json()
        
        updates = []
        params = []
        
        if '门店列表' in data:
            updates.append("门店列表 = ?")
            params.append(data['门店列表'])
        if '状态' in data:
            updates.append("状态 = ?")
            params.append(data['状态'])
        
        if updates:
            params.append(task_id)
            cursor.execute(f"""
                UPDATE 跟进任务
                SET {', '.join(updates)}
                WHERE 任务ID = ?
            """, params)
            conn.commit()
        
        conn.close()
        return jsonify({'success': True, 'message': '更新成功'})
    
    elif request.method == 'DELETE':
        cursor.execute("DELETE FROM 跟进记录 WHERE 任务ID = ?", (task_id,))
        cursor.execute("DELETE FROM 跟进任务 WHERE 任务ID = ?", (task_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': '删除成功'})

@app.route('/api/follow_tasks/<int:task_id>/stores', methods=['POST'])
def set_task_stores(task_id):
    """设置任务的门店列表"""
    data = request.get_json()
    store_codes = data.get('storeCodes', [])
    
    # 去重门店编号
    seen = set()
    unique_store_codes = []
    for code in store_codes:
        if code and code not in seen:
            seen.add(code)
            unique_store_codes.append(code)
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT 门店列表 FROM 跟进任务 WHERE 任务ID = ?", (task_id,))
    task = cursor.fetchone()
    
    if task:
        store_list = ','.join(unique_store_codes)
        cursor.execute("UPDATE 跟进任务 SET 门店列表 = ? WHERE 任务ID = ?", (store_list, task_id))
        conn.commit()
    
    conn.close()
    return jsonify({'success': True, 'message': '保存成功'})

@app.route('/api/follow_tasks/<int:task_id>/store/<store_code>/history')
def get_store_history(task_id, store_code):
    """获取门店的跟进历史记录"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT 日报数据日期, 跟进原因, 备注, 操作人, 创建时间
        FROM 跟进记录
        WHERE 任务ID = ? AND 店编号 = ?
        ORDER BY 日报数据日期 DESC, 创建时间 DESC
    """, (task_id, store_code))
    
    records = []
    for row in cursor.fetchall():
        records.append({
            '跟进日期': row['日报数据日期'],
            '跟进原因': row['跟进原因'],
            '备注': row['备注'],
            '操作人': row['操作人'],
            '创建时间': row['创建时间']
        })
    
    conn.close()
    return jsonify({'success': True, 'data': records})

@app.route('/api/follow_tasks/<int:task_id>/report')
def get_task_report(task_id):
    """生成任务报告"""
    return generate_weekly_report(task_id)

@app.route('/api/follow_records', methods=['GET', 'POST'])
def manage_follow_records():
    """获取或创建跟进记录"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if request.method == 'GET':
        task_id = request.args.get('task_id')
        date = request.args.get('date')
        store_code = request.args.get('store_code')
        
        if not task_id:
            conn.close()
            return jsonify({'success': False, 'message': '缺少任务ID'})
        
        # 先获取任务的门店列表
        cursor.execute("""
            SELECT 门店列表 FROM 跟进任务 WHERE 任务ID = ?
        """, (task_id,))
        task = cursor.fetchone()
        task_stores_str = task['门店列表'] if task else ''
        
        # 获取该日期所有门店的日报数据（去重）
        if date:
            cursor.execute("""
                SELECT 店编号, 店简称, 大区, 战区,
                       COALESCE(MAX(CAST("线索量-本地" AS REAL)), 0) as "线索量-本地",
                       COALESCE(MAX(CAST(到店数 AS REAL)), 0) as 到店数
                FROM 日报快照
                WHERE 日报数据日期 = ?
                GROUP BY 店编号, 店简称, 大区, 战区
                ORDER BY 店编号
            """, (date,))
        else:
            # 如果没有日期，获取最近日期的日报
            cursor.execute("""
                SELECT DISTINCT 日报数据日期 FROM 日报快照 ORDER BY 日报数据日期 DESC LIMIT 1
            """)
            latest_date_row = cursor.fetchone()
            latest_date = latest_date_row['日报数据日期'] if latest_date_row else None
            
            if latest_date:
                cursor.execute("""
                    SELECT 店编号, 店简称, 大区, 战区,
                           COALESCE(MAX(CAST("线索量-本地" AS REAL)), 0) as "线索量-本地",
                           COALESCE(MAX(CAST(到店数 AS REAL)), 0) as 到店数
                    FROM 日报快照
                    WHERE 日报数据日期 = ?
                    GROUP BY 店编号, 店简称, 大区, 战区
                    ORDER BY 店编号
                """, (latest_date,))
            else:
                conn.close()
                return jsonify({'success': True, 'data': []})
        
        daily_stores = [dict(row) for row in cursor.fetchall()]
        
        # 解析任务的门店列表（去重）
        task_store_codes = []
        if task_stores_str:
            try:
                import json
                store_codes = json.loads(task_stores_str)
                if isinstance(store_codes, list):
                    if store_codes and isinstance(store_codes[0], dict):
                        # 从对象数组中提取店编号并去重
                        seen = set()
                        for s in store_codes:
                            code = s.get('店编号')
                            if code and code not in seen:
                                seen.add(code)
                                task_store_codes.append(code)
                    else:
                        # 从字符串列表中去重
                        seen = set()
                        for code in store_codes:
                            if code and code not in seen:
                                seen.add(code)
                                task_store_codes.append(code)
            except:
                # 从逗号分隔字符串中去重
                codes = str(task_stores_str).split(',')
                seen = set()
                for code in codes:
                    if code and code not in seen:
                        seen.add(code)
                        task_store_codes.append(code)
        
        # 筛选任务中的门店（去重）
        target_stores = []
        if task_store_codes:
            seen = set()
            for store in daily_stores:
                code = store['店编号']
                if code in task_store_codes and code not in seen:
                    seen.add(code)
                    target_stores.append(store)
        else:
            # 如果没有任务门店列表，使用日报数据（已去重）
            target_stores = daily_stores
        
        # 获取跟进记录
        sql = """
            SELECT 记录ID, 任务ID, 日报数据日期, 店编号, 店简称,
                   "线索量-本地", 到店数, 跟进原因, 备注, 操作人, 创建时间
            FROM 跟进记录
            WHERE 任务ID = ?
        """
        params = [task_id]
        
        if date:
            sql += " AND 日报数据日期 = ?"
            params.append(date)
        
        if store_code:
            sql += " AND 店编号 = ?"
            params.append(store_code)
        
        sql += " ORDER BY 日报数据日期 DESC, 创建时间 DESC"
        
        cursor.execute(sql, params)
        records = [dict(row) for row in cursor.fetchall()]
        
        # 创建跟进记录的字典索引
        record_dict = {}
        for rec in records:
            record_dict[rec['店编号']] = rec
        
        # 合并数据
        result = []
        for store in target_stores:
            store_code = store['店编号']
            record = record_dict.get(store_code)
            
            try:
                线索量 = float(store['线索量-本地'])
                到店数 = float(store['到店数'])
                if 线索量 > 0:
                    到店率 = round(到店数 / 线索量 * 100, 2)
                else:
                    到店率 = 0
            except:
                线索量 = 0
                到店数 = 0
                到店率 = 0
            
            result.append({
                '店编号': store_code,
                '店简称': store['店简称'],
                '大区': store['大区'],
                '线索量本地': 线索量,
                '到店数': 到店数,
                '到店率': 到店率,
                '跟进原因': record['跟进原因'] if record else None,
                '备注': record['备注'] if record else None,
                '已跟进': record is not None,
                '跟进记录ID': record['记录ID'] if record else None,
                '操作人': record['操作人'] if record else None
            })
        
        conn.close()
        return jsonify({'success': True, 'data': result})
    
    elif request.method == 'POST':
        data = request.get_json()
        task_id = data.get('任务ID')
        store_code = data.get('店编号')
        date = data.get('日报数据日期')
        
        if not all([task_id, store_code, date]):
            conn.close()
            return jsonify({'success': False, 'message': '缺少必要参数'})
        
        cursor.execute("""
            INSERT INTO 跟进记录 
            (任务ID, 日报数据日期, 店编号, 跟进原因, 备注, 操作人, 创建时间)
            VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
        """, (
            task_id, date, store_code,
            data.get('跟进原因', ''),
            data.get('备注', ''),
            data.get('操作人', '')
        ))
        
        record_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': '保存成功', '记录ID': record_id})

@app.route('/api/follow_record', methods=['POST', 'PUT'])
@app.route('/api/follow_records/<int:record_id>', methods=['PUT', 'DELETE'])
def manage_follow_record(record_id=None):
    """单个跟进记录的操作"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if request.method == 'POST' or request.method == 'PUT':
        data = request.get_json()
        
        # 兼容不同的字段名
        task_id = data.get('taskId') or data.get('任务ID')
        store_code = data.get('storeCode') or data.get('店编号')
        follow_date = data.get('date') or data.get('日报数据日期')
        reason = data.get('reason') or data.get('跟进原因')
        remark = data.get('remark') or data.get('备注')
        operator = data.get('operator') or data.get('操作人')
        
        # 如果是PUT且有record_id，更新现有记录
        if record_id or data.get('记录ID'):
            rec_id = record_id or data.get('记录ID')
            
            updates = []
            params = []
            
            if reason:
                updates.append("跟进原因 = ?")
                params.append(reason)
            if remark:
                updates.append("备注 = ?")
                params.append(remark)
            if operator:
                updates.append("操作人 = ?")
                params.append(operator)
            
            if updates:
                params.append(rec_id)
                cursor.execute(f"""
                    UPDATE 跟进记录
                    SET {', '.join(updates)}
                    WHERE 记录ID = ?
                """, params)
                conn.commit()
            
            conn.close()
            return jsonify({'success': True, 'message': '更新成功', 'data': {'记录ID': rec_id}})
        else:
            # 创建新记录
            if not all([task_id, store_code, follow_date]):
                conn.close()
                return jsonify({'success': False, 'message': '缺少必要参数'})
            
            # 获取门店信息
            cursor.execute("""
                SELECT 店简称, COALESCE("线索量-本地", 0) as "线索量-本地", COALESCE(到店数, 0) as 到店数
                FROM 日报快照
                WHERE 店编号 = ? AND 日报数据日期 = ?
                LIMIT 1
            """, (store_code, follow_date))
            store_info = cursor.fetchone()
            
            店简称 = store_info['店简称'] if store_info else ''
            线索量本地 = store_info['线索量-本地'] if store_info else 0
            到店数 = store_info['到店数'] if store_info else 0
            
            cursor.execute("""
                INSERT INTO 跟进记录 
                (任务ID, 日报数据日期, 店编号, 店简称, "线索量-本地", 到店数, 跟进原因, 备注, 操作人, 创建时间)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
            """, (task_id, follow_date, store_code, 店简称, 线索量本地, 到店数, reason, remark, operator))
            
            record_id = cursor.lastrowid
            conn.commit()
            conn.close()
            return jsonify({'success': True, 'message': '保存成功', 'data': {'记录ID': record_id}})
    
    elif request.method == 'DELETE':
        cursor.execute("DELETE FROM 跟进记录 WHERE 记录ID = ?", (record_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': '删除成功'})

@app.route('/api/weekly_report/<int:task_id>')
def generate_weekly_report(task_id):
    """生成周统计报告"""
    return _generate_weekly_report(task_id)


def _generate_weekly_report(task_id):
    """生成周统计报告的内部函数"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT 任务ID, 周开始日期, 门店列表, 创建时间
        FROM 跟进任务
        WHERE 任务ID = ?
    """, (task_id,))
    task = cursor.fetchone()
    
    if not task:
        conn.close()
        return jsonify({'success': False, 'message': '任务不存在'})
    
    task_data = dict(task)
    
    # 解析门店列表
    stores = []
    store_list_str = task_data['门店列表']
    if store_list_str:
        try:
            import json
            store_codes = json.loads(store_list_str)
            if isinstance(store_codes, list):
                if store_codes and isinstance(store_codes[0], dict):
                    stores = store_codes
                else:
                    # 获取门店详情
                    codes = store_codes
                    placeholders = ','.join(['?' for _ in codes])
                    cursor.execute(f"""
                        SELECT DISTINCT 店编号, 店简称, 大区, 战区
                        FROM 日报快照
                        WHERE 店编号 IN ({placeholders})
                    """, codes)
                    stores = [dict(row) for row in cursor.fetchall()]
        except:
            codes = str(store_list_str).split(',')
            if codes and codes[0]:
                placeholders = ','.join(['?' for _ in codes])
                cursor.execute(f"""
                    SELECT DISTINCT 店编号, 店简称, 大区, 战区
                    FROM 日报快照
                    WHERE 店编号 IN ({placeholders})
                """, codes)
                stores = [dict(row) for row in cursor.fetchall()]
    
    cursor.execute("""
        SELECT DISTINCT 日报数据日期
        FROM 跟进记录
        WHERE 任务ID = ?
        ORDER BY 日报数据日期
    """, (task_id,))
    dates = [row['日报数据日期'] for row in cursor.fetchall()]
    
    cursor.execute("""
        SELECT 日报数据日期, 店编号, 跟进原因, 备注, 操作人
        FROM 跟进记录
        WHERE 任务ID = ?
        ORDER BY 日报数据日期, 店编号
    """, (task_id,))
    records = [dict(row) for row in cursor.fetchall()]
    
    conn.close()
    
    report_lines = []
    report_lines.append("=" * 60)
    report_lines.append(f"门店跟进治理周统计报告")
    report_lines.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_lines.append("=" * 60)
    report_lines.append("")
    report_lines.append(f"周次: {task_data['周开始日期']}")
    report_lines.append(f"监测门店数: {len(stores)}")
    report_lines.append(f"跟进日期: {', '.join(dates) if dates else '暂无跟进记录'}")
    report_lines.append("")
    report_lines.append("-" * 60)
    report_lines.append("各门店跟进情况:")
    report_lines.append("-" * 60)
    
    for store in stores:
        store_code = store.get('店编号', '')
        store_name = store.get('店简称', '')
        store_records = [r for r in records if r['店编号'] == store_code]
        report_lines.append(f"\n【{store_code}】{store_name}")
        
        if store_records:
            for rec in store_records:
                reason = rec['跟进原因'] or '未填写'
                remark = f"（备注：{rec['备注']}）" if rec['备注'] else ''
                operator = f" 操作人：{rec['操作人']}" if rec.get('操作人') else ''
                report_lines.append(f"  {rec['日报数据日期']}: {reason}{remark}{operator}")
        else:
            report_lines.append("  暂无跟进记录")
    
    if dates:
        report_lines.append("")
        report_lines.append("-" * 60)
        report_lines.append("跟进原因统计:")
        report_lines.append("-" * 60)
        
        all_reasons = []
        for rec in records:
            if rec['跟进原因']:
                reasons = rec['跟进原因'].split(';')
                for r in reasons:
                    r = r.strip()
                    if r:
                        all_reasons.append(r)
        
        from collections import Counter
        reason_counts = Counter(all_reasons)
        for reason, count in reason_counts.most_common():
            report_lines.append(f"  {reason}: {count}次")
    
    report_lines.append("")
    report_lines.append("=" * 60)
    report_lines.append("报告结束")
    report_lines.append("=" * 60)
    
    return jsonify({
        'success': True,
        'data': '\n'.join(report_lines)
    })

def log_operation(user_name, page, operation_type, details=None, related_id=None):
    """记录操作日志"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT 用户ID FROM 用户表 WHERE 用户名 = ?
    """, (user_name,))
    user = cursor.fetchone()
    user_id = user['用户ID'] if user else None
    
    cursor.execute("""
        INSERT INTO 操作日志 (用户ID, 用户名, 操作页面, 操作类型, 操作详情, 关联ID, IP地址, 操作时间)
        VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
    """, (
        user_id,
        user_name,
        page,
        operation_type,
        json.dumps(details, ensure_ascii=False) if details else None,
        related_id,
        request.remote_addr
    ))
    
    conn.commit()
    conn.close()

@app.route('/api/users', methods=['GET', 'POST'])
def manage_users():
    """用户管理"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if request.method == 'GET':
        cursor.execute("""
            SELECT 用户ID, 用户名, 显示名称, 最后登录时间
            FROM 用户表
            ORDER BY 最后登录时间 DESC
        """)
        users = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return jsonify({'success': True, 'data': users})
    
    elif request.method == 'POST':
        data = request.get_json()
        username = data.get('用户名', '').strip()
        
        if not username:
            conn.close()
            return jsonify({'success': False, 'message': '用户名不能为空'})
        
        cursor.execute("SELECT * FROM 用户表 WHERE 用户名 = ?", (username,))
        if cursor.fetchone():
            cursor.execute("""
                UPDATE 用户表 SET 最后登录时间 = datetime('now')
                WHERE 用户名 = ?
            """, (username,))
        else:
            cursor.execute("""
                INSERT INTO 用户表 (用户名, 显示名称, 最后登录时间, 创建时间)
                VALUES (?, ?, datetime('now'), datetime('now'))
            """, (username, username))
        
        conn.commit()
        conn.close()
        
        log_operation(username, '系统', '登录')
        
        return jsonify({'success': True, 'message': '登录成功', '用户名': username})

@app.route('/api/operation_logs', methods=['GET', 'POST'])
def manage_operation_logs():
    """操作日志管理"""
    if request.method == 'GET':
        user_name = request.args.get('user_name')
        page = request.args.get('page')
        operation_type = request.args.get('operation_type')
        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')
        limit = request.args.get('limit', 100, type=int)
        
        conn = get_db_connection()
        cursor = conn.cursor()
        
        sql = """
            SELECT 日志ID, 用户名, 操作页面, 操作类型, 操作详情, 关联ID, IP地址, 操作时间
            FROM 操作日志
            WHERE 1=1
        """
        params = []
        
        if user_name:
            sql += " AND 用户名 = ?"
            params.append(user_name)
        
        if page:
            sql += " AND 操作页面 = ?"
            params.append(page)
        
        if operation_type:
            sql += " AND 操作类型 = ?"
            params.append(operation_type)
        
        if start_date:
            sql += " AND date(操作时间) >= date(?)"
            params.append(start_date)
        
        if end_date:
            sql += " AND date(操作时间) <= date(?)"
            params.append(end_date)
        
        sql += " ORDER BY 操作时间 DESC LIMIT ?"
        params.append(limit)
        
        cursor.execute(sql, params)
        logs = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return jsonify({'success': True, 'data': logs})
    
    elif request.method == 'POST':
        data = request.get_json()
        user_name = data.get('用户名')
        page = data.get('操作页面')
        operation_type = data.get('操作类型')
        details = data.get('操作详情')
        related_id = data.get('关联ID')
        
        if not all([user_name, page, operation_type]):
            return jsonify({'success': False, 'message': '缺少必要参数'})
        
        log_operation(user_name, page, operation_type, details, related_id)
        
        return jsonify({'success': True, 'message': '日志记录成功'})

@app.route('/operation_logs')
def operation_logs_page():
    """操作日志页面"""
    return render_template('operation_logs.html')

@app.route('/store_profile')
def store_profile_page():
    """门店档案页面"""
    return render_template('store_profile.html')

@app.route('/store_detail/<store_code>')
def store_detail_page(store_code):
    """门店详情页面"""
    return render_template('store_detail.html')

@app.route('/api/store_profile/search')
def search_stores():
    """搜索门店"""
    keyword = request.args.get('q', '')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if keyword:
        cursor.execute("""
            SELECT DISTINCT 店编号, 店简称, 大区, 战区
            FROM 日报快照
            WHERE 店编号 LIKE ? OR 店简称 LIKE ?
            ORDER BY 店编号
            LIMIT 20
        """, (f'%{keyword}%', f'%{keyword}%'))
    else:
        cursor.execute("""
            SELECT DISTINCT 店编号, 店简称, 大区, 战区
            FROM 日报快照
            ORDER BY 店编号
            LIMIT 20
        """)
    
    stores = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    return jsonify({'success': True, 'data': stores})

@app.route('/api/store_profile/<store_code>/basic_info')
def get_store_basic_info(store_code):
    """获取门店基础信息"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT DISTINCT 店编号, 店简称, 大区, 战区, 大区经理, 战区经理
        FROM 日报快照
        WHERE 店编号 = ?
        LIMIT 1
    """, (store_code,))
    
    store = cursor.fetchone()
    conn.close()
    
    if not store:
        return jsonify({'success': False, 'message': '门店不存在'})
    
    return jsonify({'success': True, 'data': dict(store)})

@app.route('/api/store_profile/<store_code>/daily_stats')
def get_store_daily_stats(store_code):
    """获取门店日报统计数据"""
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT 日报数据日期, "线索量-本地", 到店数
        FROM 日报快照
        WHERE 店编号 = ?
        ORDER BY 日报数据日期
    """, (store_code,))
    
    rows = cursor.fetchall()
    conn.close()
    
    data = []
    for row in rows:
        date_str = row['日报数据日期']
        converted_date = convert_date_format(date_str)
        
        in_range = True
        if start_date and converted_date < start_date:
            in_range = False
        if end_date and converted_date > end_date:
            in_range = False
        
        if in_range:
            线索量 = float(row['线索量-本地']) if row['线索量-本地'] else 0
            到店数 = float(row['到店数']) if row['到店数'] else 0
            到店率 = round(到店数 / 线索量 * 100, 2) if 线索量 > 0 else 0
            
            data.append({
                '日报数据日期': row['日报数据日期'],
                '线索量数值': 线索量,
                '到店数数值': 到店数,
                '到店率': 到店率
            })
    
    return jsonify({'success': True, 'data': data})

@app.route('/api/store_profile/<store_code>/follow_history')
def get_store_follow_history(store_code):
    """获取门店历史跟进记录"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT r.记录ID, r.任务ID, r.日报数据日期, r.跟进原因, r.备注, r.操作人, r.创建时间,
               t.周开始日期
        FROM 跟进记录 r
        LEFT JOIN 跟进任务 t ON r.任务ID = t.任务ID
        WHERE r.店编号 = ?
        ORDER BY r.创建时间 DESC
    """, (store_code,))
    
    rows = cursor.fetchall()
    conn.close()
    
    records = [dict(row) for row in rows]
    
    return jsonify({'success': True, 'data': records})

@app.route('/api/store_profile/follow_summary')
def get_follow_summary():
    """获取跟进汇总统计"""
    min_times = request.args.get('min_times', 1, type=int)
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT COUNT(DISTINCT 店编号) as total_stores,
               COUNT(*) as total_records
        FROM 跟进记录
    """)
    total = cursor.fetchone()
    
    total_stores = total['total_stores'] or 0
    total_records = total['total_records'] or 0
    avg_times = round(total_records / total_stores, 2) if total_stores > 0 else 0
    
    cursor.execute("""
        SELECT COUNT(DISTINCT 店编号) as high_freq_stores
        FROM (
            SELECT 店编号, COUNT(*) as cnt
            FROM 跟进记录
            GROUP BY 店编号
            HAVING COUNT(*) >= 3
        )
    """)
    high_freq = cursor.fetchone()
    
    cursor.execute("""
        SELECT 跟进原因, COUNT(*) as 出现次数,
               COUNT(DISTINCT 店编号) as 涉及门店数
        FROM 跟进记录
        WHERE 跟进原因 IS NOT NULL AND 跟进原因 != ''
        GROUP BY 跟进原因
        ORDER BY 出现次数 DESC
        LIMIT 10
    """)
    reason_stats = [dict(row) for row in cursor.fetchall()]
    
    for reason in reason_stats:
        reason['占比'] = round(reason['出现次数'] * 100 / total_records, 2) if total_records > 0 else 0
    
    conn.close()
    
    return jsonify({
        'success': True,
        'data': {
            'total_stores': total_stores,
            'total_records': total_records,
            'avg_times': avg_times,
            'high_freq_stores': high_freq['high_freq_stores'] or 0,
            'reason_stats': reason_stats
        }
    })

@app.route('/api/store_profile/regions')
def get_regions():
    """获取大区列表"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT DISTINCT 大区
        FROM 日报快照
        WHERE 大区 IS NOT NULL AND 大区 != ''
        ORDER BY 大区
    """)
    
    regions = [row['大区'] for row in cursor.fetchall()]
    conn.close()
    
    return jsonify({'success': True, 'data': {'regions': regions}})

@app.route('/api/store_profile/war_zones')
def get_war_zones():
    """获取战区列表（可选按大区筛选）"""
    region = request.args.get('region', '')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if region:
        cursor.execute("""
            SELECT DISTINCT 战区
            FROM 日报快照
            WHERE 大区 = ? AND 战区 IS NOT NULL AND 战区 != ''
            ORDER BY 战区
        """, (region,))
    else:
        cursor.execute("""
            SELECT DISTINCT 战区
            FROM 日报快照
            WHERE 战区 IS NOT NULL AND 战区 != ''
            ORDER BY 战区
        """)
    
    war_zones = [row['战区'] for row in cursor.fetchall()]
    conn.close()
    
    return jsonify({'success': True, 'data': {'war_zones': war_zones}})

def convert_date_format(date_str):
    """将日期格式从 'YYYY/M/D' 转换为 'YYYY-MM-DD'"""
    try:
        parts = date_str.split('/')
        if len(parts) == 3:
            return f"{parts[0]}-{int(parts[1]):02d}-{int(parts[2]):02d}"
        return date_str
    except:
        return date_str

@app.route('/api/store_profile/store_list')
def get_store_list():
    """获取门店列表（支持时间段筛选、分页）"""
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')
    region = request.args.get('region', '')
    war_zone = request.args.get('war_zone', '')
    search = request.args.get('search', '')
    follow_status = request.args.get('follow_status', '')
    page = request.args.get('page', 1, type=int)
    page_size = request.args.get('page_size', 20, type=int)
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    base_filter = ""
    filter_params = []
    
    if region:
        base_filter += " AND 大区 = ?"
        filter_params.append(region)
    
    if war_zone:
        base_filter += " AND 战区 = ?"
        filter_params.append(war_zone)
    
    if search:
        base_filter += " AND (店编号 LIKE ? OR 店简称 LIKE ?)"
        filter_params.append(f'%{search}%')
        filter_params.append(f'%{search}%')
    
    total_sql = f"""
        SELECT COUNT(DISTINCT s.店编号)
        FROM 日报快照 s
        LEFT JOIN (SELECT DISTINCT 店编号 FROM 跟进记录) f ON s.店编号 = f.店编号
        WHERE 1=1 {base_filter.replace('店编号', 's.店编号').replace('店简称', 's.店简称').replace('大区', 's.大区').replace('战区', 's.战区')}
    """
    
    if follow_status == 'followed':
        total_sql += " AND f.店编号 IS NOT NULL"
    elif follow_status == 'unfollowed':
        total_sql += " AND f.店编号 IS NULL"
    
    cursor.execute(total_sql, filter_params)
    total = cursor.fetchone()
    total_count = total[0] if total else 0
    
    cursor.execute(f"""
        SELECT DISTINCT 店编号, 店简称, 大区, 战区
        FROM 日报快照
        WHERE 1=1 {base_filter}
        ORDER BY 店编号
        LIMIT ? OFFSET ?
    """, filter_params + [page_size, (page - 1) * page_size])
    
    store_rows = cursor.fetchall()
    
    store_codes = [row['店编号'] for row in store_rows]
    
    if store_codes and (start_date or end_date):
        cursor.execute("""
            SELECT 店编号, 日报数据日期, "线索量-本地", 到店数
            FROM 日报快照
            WHERE 店编号 IN ({})
        """.format(','.join(['?' for _ in store_codes])), store_codes)
        
        date_data = {}
        for row in cursor.fetchall():
            date_str = row['日报数据日期']
            converted_date = convert_date_format(date_str)
            
            in_range = True
            if start_date and converted_date < start_date:
                in_range = False
            if end_date and converted_date > end_date:
                in_range = False
            
            if in_range:
                code = row['店编号']
                clue = float(row['线索量-本地']) if row['线索量-本地'] else 0
                arrival = float(row['到店数']) if row['到店数'] else 0
                
                if code not in date_data:
                    date_data[code] = {'clues': [], 'arrivals': []}
                date_data[code]['clues'].append(clue)
                date_data[code]['arrivals'].append(arrival)
        
        for code in date_data:
            clues = date_data[code]['clues']
            arrivals = date_data[code]['arrivals']
            date_data[code]['avg_clue'] = sum(clues) / len(clues) if clues else 0
            date_data[code]['avg_arrival'] = sum(arrivals) / len(arrivals) if arrivals else 0
    else:
        cursor.execute("""
            SELECT 店编号,
                   AVG(CAST("线索量-本地" AS REAL)) as avg_clue,
                   AVG(CAST(到店数 AS REAL)) as avg_arrival
            FROM 日报快照
            WHERE 店编号 IN ({})
            GROUP BY 店编号
        """.format(','.join(['?' for _ in store_codes])), store_codes)
        
        date_data = {}
        for row in cursor.fetchall():
            date_data[row['店编号']] = {
                'avg_clue': float(row['avg_clue']) if row['avg_clue'] else 0,
                'avg_arrival': float(row['avg_arrival']) if row['avg_arrival'] else 0
            }
    
    cursor.execute("""
        SELECT DISTINCT 店编号 FROM 跟进记录
    """)
    followed_stores = set([row['店编号'] for row in cursor.fetchall()])
    
    conn.close()
    
    stores = []
    for row in store_rows:
        code = row['店编号']
        data = date_data.get(code, {'avg_clue': 0, 'avg_arrival': 0})
        
        avg_clue = data['avg_clue']
        avg_arrival = data['avg_arrival']
        avg_rate = round(avg_arrival / avg_clue * 100, 2) if avg_clue > 0 else 0
        
        stores.append({
            '店编号': row['店编号'],
            '店简称': row['店简称'],
            '大区': row['大区'],
            '战区': row['战区'],
            '平均线索量': round(avg_clue, 0),
            '平均到店数': round(avg_arrival, 0),
            '平均到店率': avg_rate,
            '是否跟进': code in followed_stores
        })
    
    total_pages = (total_count + page_size - 1) // page_size
    
    return jsonify({
        'success': True,
        'data': {
            'list': stores,
            'pagination': {
                'page': page,
                'page_size': page_size,
                'total': total_count,
                'total_pages': total_pages
            }
        }
    })

@app.route('/api/store_profile/frequent_stores')
def get_frequent_stores():
    """获取被多次跟进的门店列表"""
    min_times = request.args.get('min_times', 1, type=int)
    region = request.args.get('region', '')
    war_zone = request.args.get('war_zone', '')
    search = request.args.get('search', '')
    sort = request.args.get('sort', 'follow_times')
    order = request.args.get('order', 'desc')
    page = request.args.get('page', 1, type=int)
    page_size = request.args.get('page_size', 20, type=int)
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT r.店编号,
               COUNT(*) as 跟进次数,
               MIN(r.创建时间) as 首次跟进时间,
               MAX(r.创建时间) as 最后跟进时间,
               GROUP_CONCAT(DISTINCT r.跟进原因) as 跟进原因_summary
        FROM 跟进记录 r
        GROUP BY r.店编号
        HAVING COUNT(*) >= ?
        ORDER BY 跟进次数 DESC
    """, (min_times,))
    
    store_stats = [dict(row) for row in cursor.fetchall()]
    
    store_codes = [s['店编号'] for s in store_stats]
    
    if not store_codes:
        conn.close()
        return jsonify({
            'success': True,
            'data': [],
            'pagination': {
                'page': page,
                'page_size': page_size,
                'total': 0,
                'total_pages': 0
            }
        })
    
    placeholders = ','.join(['?'] * len(store_codes))
    cursor.execute(f"""
        SELECT DISTINCT 店编号, 店简称, 大区, 战区
        FROM 日报快照
        WHERE 店编号 IN ({placeholders})
    """, store_codes)
    
    store_info = {row['店编号']: dict(row) for row in cursor.fetchall()}
    
    result = []
    for store in store_stats:
        code = store['店编号']
        if code in store_info:
            info = store_info[code]
            
            if region and info.get('大区') != region:
                continue
            if war_zone and info.get('战区') != war_zone:
                continue
            if search and (code.lower().find(search.lower()) == -1 and info.get('店简称', '').lower().find(search.lower()) == -1):
                continue
            
            result.append({
                '店编号': code,
                '店简称': info.get('店简称', ''),
                '大区': info.get('大区', ''),
                '战区': info.get('战区', ''),
                '跟进次数': store['跟进次数'],
                '首次跟进时间': store['首次跟进时间'],
                '最后跟进时间': store['最后跟进时间'],
                '跟进原因_summary': store.get('跟进原因_summary', '')
            })
    
    if sort == 'follow_times':
        result.sort(key=lambda x: x['跟进次数'], reverse=(order == 'desc'))
    elif sort == 'store_code':
        result.sort(key=lambda x: x['店编号'], reverse=(order == 'desc'))
    elif sort == 'region':
        result.sort(key=lambda x: x.get('大区', ''), reverse=(order == 'desc'))
    
    total = len(result)
    total_pages = (total + page_size - 1) // page_size
    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    result = result[start_idx:end_idx]
    
    conn.close()
    
    return jsonify({
        'success': True,
        'data': result,
        'pagination': {
            'page': page,
            'page_size': page_size,
            'total': total,
            'total_pages': total_pages
        }
    })

@app.route('/api/store_profile/reason_analysis')
def get_reason_analysis():
    """获取原因分析"""
    min_times = request.args.get('min_times', 1, type=int)
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT COUNT(*) as total_records
        FROM 跟进记录
    """)
    total_records = cursor.fetchone()['total_records'] or 0
    
    cursor.execute("""
        SELECT 跟进原因, COUNT(*) as 出现次数,
               COUNT(DISTINCT 店编号) as 涉及门店数
        FROM 跟进记录
        WHERE 跟进原因 IS NOT NULL AND 跟进原因 != ''
        GROUP BY 跟进原因
        ORDER BY 出现次数 DESC
    """)
    reasons = [dict(row) for row in cursor.fetchall()]
    
    for reason in reasons:
        reason['占比'] = round(reason['出现次数'] * 100 / total_records, 2) if total_records > 0 else 0
        
        cursor.execute("""
            SELECT DISTINCT 店编号
            FROM 跟进记录
            WHERE 跟进原因 LIKE '%' || ? || '%'
        """, (reason['跟进原因'],))
        stores = [row['店编号'] for row in cursor.fetchall()]
        reason['门店列表'] = stores
    
    conn.close()
    
    return jsonify({
        'success': True,
        'data': {
            'reason_distribution': reasons
        }
    })

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5003)

