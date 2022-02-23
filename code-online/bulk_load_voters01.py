import pymysql
from bulk_load_functions import csv_to_mysql
import pandas as pd
import sys


host = 'localhost'
user = 'root'
password = 'Baci8888'
database = 'ddnevada'  

def main():
    
    # Execution Example
    #sqlQuery = "CREATE TABLE Voters(id int, LastName varchar(32), FirstName varchar(32), DepartmentCode int)"   

    db_delete = "DROP DATABASE IF EXISTS " + database;
    db_create = "CREATE DATABASE IF NOT EXISTS "  + database;

    load_local  = "SET GLOBAL local_infile=1"
    base_delete = "DROP TABLE IF EXISTS " + database + ".voters"
    hist_delete = "DROP TABLE IF EXISTS " + database + ".history"
    base_create = "CREATE TABLE " + database + ".voters(countyid int, stateid int, status varchar(8), county int, precinct int, \
        congdist int, assmDist int, senDist int, brdofed int, regent int, cntycomm int, rwards int, swards int, schbtrust char, schbdatlrg char, \
        first varchar(30), last varchar(30), middle varchar(30), suffix varchar(8), phone varchar(20), \
        email varchar(50), birthDate varchar(8), regdate varchar(12), party varchar(30), \
        streetno varchar(12), streetname varchar(30), address1 varchar(30),	address2 varchar(30), city varchar(30),	state varchar(4), zip varchar(10), \
        daysregistered int, age	int, \
        11_03_20_general varchar(8), 06_09_20_primary varchar(8), 11_06_18_general varchar(8), 06_12_18_primary varchar(8), \
        11_08_16_general varchar(8), 06_14_16_primary varchar(8), 11_04_14_general varchar(8), 06_10_14_primary varchar(8), \
        11_06_12_general varchar(8), 06_12_12_primary varchar(8), 11_02_10_general varchar(8), 06_08_10_primary varchar(8), \
        11_04_08_general varchar(8), 08_12_08_primary varchar(8), 11_07_06_general varchar(8), 08_15_06_primary varchar(8), \
        11_02_04_general varchar(8), 09_07_04_primary varchar(8), 11_05_02_general varchar(8), 09_03_02_primary varchar(8), \
        totalvotes int,	generals int, primaries int, polls int,	absentee int, early int, provisional int, \
        likelytoVote varchar(12), score int, \
        PRIMARY KEY(stateid))"
        
    basesm_create = "CREATE TABLE " + database + ".voters(stateid int, countyid int, \
        first varchar(30), last varchar(30), middle varchar(30), \
        birthdate varchar(8), regdate varchar(12), daysregistered int, party varchar(30), \
        PRIMARY KEY(stateid))"
    hist_create = "CREATE TABLE " + database + ".history(id_history int, stateid int, voter_strength varchar(30), voter_level int, PRIMARY KEY(id_history))"
    #load_base_d = "LOAD DATA LOCAL INFILE '/Users/jimsievers/Desktop/@ONLINE-voters/db2/base_washoe.csv' INTO TABLE " + database + ".voters FIELDS TERMINATED BY ',' LINES TERMINATED BY '\n' IGNORE 1 LINES "
    #load_base_d = "LOAD DATA LOCAL INFILE '/Users/jimsievers//Users/jimsievers/Desktop/base_state.csv' INTO TABLE " + database + ".voters FIELDS TERMINATED BY ',' LINES TERMINATED BY '\n' IGNORE 1 LINES "
    load_base_d = "LOAD DATA LOCAL INFILE '/Users/jimsievers/Desktop/base_state.csv' INTO TABLE " + database + ".voters FIELDS TERMINATED BY ',' LINES TERMINATED BY '\n' IGNORE 1 LINES "
    #load_hist_d = "LOAD DATA LOCAL INFILE '/Users/jimsievers//Users/jimsievers/iCloud/my\ folders/my\ desktop/db2/vote_hist.csv' INTO TABLE " + database + ".history FIELDS TERMINATED BY ',' LINES TERMINATED BY '\n' IGNORE 1 LINES "
    
    set_password = "ALTER USER 'root'@'localhost' IDENTIFIED BY 'Baci8888'"

    
    #delete all application tables from database
            
            
    #create new application tables
    sqlcreate = "CREATE TABLE voters(id int, ci_id int, vh_id int, vi_id int)" 


            
    #load the application tables   
    # base    
    #csv_to_mysql(set_password, host, user, password)
    #csv_to_mysql(db_delete, host, user, password)
    #csv_to_mysql(db_create, host, user, password)
    #csv_to_mysql(database, db_delete, host, user, password)
    #csv_to_mysql(database, db_create, host, user, password)
    csv_to_mysql(database, load_local, host, user, password)
    csv_to_mysql(database, base_delete, host, user, password)
    csv_to_mysql(database, base_create, host, user, password)
    csv_to_mysql(database, load_base_d, host, user, password)
    # history
    #csv_to_mysql(database, hist_delete, host, user, password)
    #csv_to_mysql(database, hist_create, host, user, password)
    #csv_to_mysql(database, load_hist_d, host, user, password)

#*******************************************************************************
#  Standard boilerplate to call the main() function to begin                   *
#  the program.  This allows this script to be imported into another one       *
#  and not try to run the show in that case as __name__ will not be __main__.  *
#  When the script is run directly this will evaluate to TRUE and thus         *
#  call the function main and make things work as expected.                    *
#                                                                              *
#  Not really needed for this program, but good practice for the future.       *
#                                                                              *
# ******************************************************************************
if __name__ == '__main__':
    main()