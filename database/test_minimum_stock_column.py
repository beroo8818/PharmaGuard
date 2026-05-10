from database.schema import connect, init_database

init_database()

conn = connect()
cur = conn.cursor()

cur.execute("PRAGMA table_info(items);")
columns = [row[1] for row in cur.fetchall()]

conn.close()

if "minimum_stock" in columns:
    print("OK: minimum_stock column exists.")
else:
    print("ERROR: minimum_stock column is missing.")