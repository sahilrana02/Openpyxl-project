import mysql.connector as mdb
conn=mdb.connect(host='127.0.0.1', user='root', password='')
print(conn)

