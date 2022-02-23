import pymysql
import pandas as pd
import sys

def csv_to_mysql(database ,my_sql, host, user, password):
 
    host = 'localhost'
    user = 'root'
    password = 'Baci8888'


    try:
        con = pymysql.connect(host=host,
                                user=user,
                                password=password,
                                autocommit=True,
                                local_infile=1,
                                db=database)
        print('Connected to DB: {}'.format(host))
        # Create cursor and execute Load SQL
        cursor = con.cursor()
        
        print('EXECUTE SQL: {}'.format(my_sql))

        cursor.execute(my_sql)
        print('STATUS    :    Successfully executed SQL.')
        con.close()
       
    except Exception as e:
        print('Error: {}'.format(str(e)))
        sys.exit(1)

