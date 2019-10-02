import mysql.connector as mdb

def run_sql_file(filename, connection):
    file = open(filename, 'r')
    sql = s = " ".join(file.readlines())
    print ("Start executing sql: ")
    cursor = connection.cursor()
    cursor.execute(sql)

def main():
    connection = mdb.connect('127.0.0.1', 'root', 'password', 'database_name')
    with open(r"C:\Users\RANA\Desktop\PFA\R1 Jobs1.txt","r") as obj:
        for j in obj:
            run_sql_file(r'C:\Users\RANA\Desktop\PFA\3 New\3 New\{}\MIM\SQL\delete-JOB.sql'.format(j), connection)
            run_sql_file(r'C:\Users\RANA\Desktop\PFA\3 New\3 New\{}\MIM\SQL\insert-JOB-DEV-SYS-INT-PRD.sql'.format(j), connection)
    connection.close()

if __name__ == "__main__":
    main()
