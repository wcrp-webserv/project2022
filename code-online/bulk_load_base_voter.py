import pymysql
from bulk_load_functions import csv_to_mysql
import pandas as pd
import sys


host = 'localhost'
user = 'root'
password = 'Baci8888'
database = 'base_voterlg'  

def main():
    


    base_delete = "DROP TABLE IF EXISTS " + database + ".voters"
    addr_delete = "DROP TABLE IF EXISTS " + database + ".addr"
    base_create = "CREATE TABLE " + database + ".voters(stateid  varchar(12), countyid  varchar(12), status varchar(8), county varchar(12), precinct varchar(12), \
        first varchar(30), last varchar(30), middle varchar(30), suffix varchar(8),  \
        birthDate varchar(12), regdate varchar(12), regdays int, age int, party varchar(30), \
        totalvotes int,	generals int, primaries int, polls int,	absentee int, early int, provisional int, likelytoVote varchar(12), score int, \
        PRIMARY KEY(stateid))"
    
    addr_create = "CREATE TABLE " + database + ".addr(id_addr int, stateid int, countyid int, phone varchar(12), email varchar(45), \
        streetno varchar(12), streetname varchar(30), address1 varchar(30), address2 varchar(30), city varchar(30), state varchar(4), zip varchar(12), \
        congdist varchar(5), assmdist varchar(5), sendist varchar(5), brdofed varchar(5), regent varchar(5), cntycomm varchar(5), rwards varchar(5), swards varchar(5), schbdtrust varchar(5), schbdatlrg varchar(5), \
        PRIMARY KEY(id_addr))"
            
        #11_03_20_general varchar(8), 06_09_20_primary varchar(8), 11_06_18_general varchar(8), 06_12_18_primary varchar(8), \
        
    hist_create = "CREATE TABLE " + database + ".addr(id_addr int, stateid int, voter_strength varchar(30), voter_level int, PRIMARY KEY(id_history))"
    load_base_d = "LOAD DATA LOCAL INFILE '/Users/jimsievers/Desktop/@base_table/base_voter.csv' INTO TABLE " + database + ".voters FIELDS TERMINATED BY ',' LINES TERMINATED BY '\n' IGNORE 1 LINES "
    load_addr_d = "LOAD DATA LOCAL INFILE '/Users/jimsievers/Desktop/@base_table/addr_voter.csv' INTO TABLE " + database + ".addr FIELDS TERMINATED BY ',' LINES TERMINATED BY '\n' IGNORE 1 LINES "
    
    set_password = "ALTER USER 'root'@'localhost' IDENTIFIED BY 'Baci8888'"

    
            
    # Execution Example
    #sqlQuery = "CREATE TABLE Voters(id int, LastName varchar(32), FirstName varchar(32), DepartmentCode int)"   
    #streetno varchar(12), streetname varchar(30), address1 varchar(30),	address2 varchar(30), city varchar(30),	state varchar(4), zip varchar(10), phone varchar(20), email varchar(50), \

    #db_delete = "DROP DATABASE IF EXISTS " + database;
    #db_create = "CREATE DATABASE IF NOT EXISTS "  + database;         
    #create new application tables
    #sqlcreate = "CREATE TABLE voters(id int, ci_id int, vh_id int, vi_id int)" 


            
    #load the application tables   
    '''
    csv_to_mysql(database, base_delete, host, user, password)
    csv_to_mysql(database, base_create, host, user, password)
    csv_to_mysql(database, load_base_d, host, user, password)
    '''
    # addr
    #csv_to_mysql(database, addr_delete, host, user, password)
    #csv_to_mysql(database, addr_create, host, user, password)
    #csv_to_mysql(database, load_addr_d, host, user, password)

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