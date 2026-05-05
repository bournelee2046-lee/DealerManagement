
import sqlite3

conn = sqlite3.connect('日报.db')
cursor = conn.cursor()

cursor.execute("SELECT DISTINCT 日报数据日期 FROM 日报快照 ORDER BY 日报数据日期 DESC")
dates = cursor.fetchall()
print('数据库中的日期:')
for d in dates:
    print(f'  - {d[0]}')

if dates:
    first_date = dates[0][0]
    cursor.execute("SELECT COUNT(*) FROM 日报快照 WHERE 日报数据日期 = ?", (first_date,))
    count = cursor.fetchone()[0]
    print(f'\n日期 {first_date} 有 {count} 条记录')
    
    cursor.execute("SELECT * FROM 日报快照 WHERE 日报数据日期 = ? LIMIT 3", (first_date,))
    print('\n前3条记录:')
    for row in cursor.fetchall():
        print(row)

conn.close()
