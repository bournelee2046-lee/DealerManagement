
import sqlite3

conn = sqlite3.connect('日报.db')
cursor = conn.cursor()

cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = cursor.fetchall()
print('Tables:', tables)

if tables:
    for table in tables:
        table_name = table[0]
        print(f'\n--- {table_name} ---')
        cursor.execute(f"PRAGMA table_info({table_name})")
        for col in cursor.fetchall():
            print(col)

conn.close()
