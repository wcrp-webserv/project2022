from typing import Type
import pandas as pd
import numpy as np
import sys, os
import csv
import argparse
from os.path import isfile, join
import datetime
import time
import xlsxwriter
import math


# Constants
DIST_AD     = 6
DIST_SD     = 7
DIST_RWARDS = 11
DIST_CNTYCM = 10

# Variables
district = ""
districtid = ""
ProgName = "SELECTROWS"                 # Name of running program
out_file = "deafult_name.csv"
base_file = "/Users/jimsievers/symlinks/nvsos-nevada/state-20220212/state/base.csv"

def printLine (printData):
    global printFileh,ProgName
    datestring = datetime.datetime.now()
    datestring = datestring.strftime("%m/%d/%Y %H:%M:%S")
    if (printData[-1] == "\r"):
        print( ProgName + " " + datestring + ' - ' + printData, end="")
        return
    print( ProgName + " " + datestring + ' - ' + printData)
    print( ProgName + " " + datestring + ' - ' + printData, file=printFileh)
    return

def printhelp():
    print ("py turnout.py  -i <basefile> -r <reportfile> -p <precinct> -d <district>")
    print ("    -i <base>       = CountyProcessed Secretary of state data file.")
    print ("                      Default is base.csv.")
    print ("    -r <output>     = output report Excel Spreadsheet file")
    print ("                      Default is output.xlsx")
    print ("    -c <candidate>  = Candidate name.")
    print ("    -d <district>   = District to report turnout for (CD, AD, SD, CntyComm, Rwards")
    print (" ")
    return

#----------------------------------------------------------
# get arguments
# -d, --district, political district (SD, AD, CityCoucil, CntyComm and number SD 8)
# -b, --base, base file for voter records
# -o, --output, output file for extracted
#---------------------------------------------------------- 


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--district', help= "Select a District", nargs=2, required=True)
    parser.add_argument('-b', '--base', help= "Base file", default="base2.csv", required=True)
    parser.add_argument('-o', '--output', help= "Output file" , default="output.csv")
    parser.add_argument('-c', '--candidate', help= "Candidate Name" , default="Unknown-candidate")
    args = parser.parse_args()

#----------------------------------------------------------
# obtain argument values
#---------------------------------------------------------- 
    if args.district[0] == "SD":
        district = DIST_SD
        districtid = args.district[1]
    elif args.district[0] == "AD":
        district = DIST_AD
        districtid = args.district[1]
    elif args.district[0] == "Rwards":
        district = DIST_RWARDS
        districtid = args.district[1]
    elif args.district[0] == "CntyComm":
        district = DIST_CNTYCM
        districtid = args.district[1]
    else:
        print ("Error district")
        exit(2)
    base_file = args.base
    out_file = args.output
        
#----------------------------------------------------------
# program loop 
#----------------------------------------------------------          
    with open(base_file) as csv_in, open (out_file, 'w', newline='') as csv_out:
        csvreader = csv.reader(csv_in, delimiter=',')
        csvwriter = csv.writer(csv_out, delimiter=',')
        
        row1 = next(csvreader)
        csvwriter.writerow(row1)
        
        for row in csvreader:
            if row[district] == districtid:
                csvwriter.writerow(row)
                
    return(0)
            
#----------------------------------------------------------
#  Standard boilerplate to call the main() function to begin                    
#  the program.  This allows this script to be imported into another one       
#  and not try to run the show in that case as __name__ will not be __main__.  
#  When the script is run directly this will evaluate to TRUE and thus         
#  call the function main and make things work as expected.                    
#                                                                              
#  Not really needed for this program, but good practice for the future.       
#                                                                             
#----------------------------------------------------------
if __name__ == '__main__':
    exit (main())