
import sqlite3

def test_search():
    print('=' * 80)
    print('测试搜索功能')
    print('=' * 80)
    
    conn = sqlite3.connect('日报.db')
    cursor = conn.cursor()
    
    # 测试1：搜索店编号
    print('\n1. 测试搜索店编号（GDB0100）')
    cursor.execute("SELECT * FROM 日报快照 WHERE 日报数据日期 = '2026/4/30' AND (店编号 LIKE ? OR 店简称 LIKE ?)", ('%GDB0100%', '%GDB0100%'))
    rows = cursor.fetchall()
    print(f'找到 {len(rows)} 条记录')
    if rows:
        print(f'第一条：店编号={rows[0][7]}, 店简称={rows[0][8]}')
    
    # 测试2：搜索店简称
    print('\n2. 测试搜索店简称（深圳）')
    cursor.execute("SELECT * FROM 日报快照 WHERE 日报数据日期 = '2026/4/30' AND (店编号 LIKE ? OR 店简称 LIKE ?)", ('%深圳%', '%深圳%'))
    rows = cursor.fetchall()
    print(f'找到 {len(rows)} 条记录')
    for i, row in enumerate(rows[:3]):
        print(f'{i+1}. 店编号={row[7]}, 店简称={row[8]}')
    
    # 测试3：日期和搜索组合
    print('\n3. 测试日期筛选 + 搜索（2026/4/30 + 龙华）')
    cursor.execute("SELECT * FROM 日报快照 WHERE 日报数据日期 = '2026/4/30' AND (店编号 LIKE ? OR 店简称 LIKE ?)", ('%龙华%', '%龙华%'))
    rows = cursor.fetchall()
    print(f'找到 {len(rows)} 条记录')
    
    conn.close()
    
    print('\n' + '=' * 80)
    print('搜索功能测试完成！')
    print('=' * 80)

if __name__ == '__main__':
    test_search()

