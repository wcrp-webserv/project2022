
#-----------------------------------------------------------------------------------
#                               voter_screen.py                                     
#                                                                                   
#  purpose: currently creates some statistics by analyzing a base.csv type file and
#   input:  analyze a file in the form of base.csv from nvvoter2.pl
#   output: various counts and averages from that file
#                
#----------------------------------------------------------------------------------- 
import numpy as np
import argparse
import errno, os
from csv  import reader
from csv import DictReader


#
# global variables
# [[ should only nave to be specified once in this program file ]]
# [[ now i must delclate them a global again in any function ]]
#

rownum = 0
allvoters = 0

democrats = 0
republicans = 0
other = 0
nonps = 0
iaps = 0
anyvote = 0
repvote = 0
demvote = 0
othvote = 0
nonpvote = 0
iapvote = 0
vote = ""
voted = ['BR', 'FW', 'EV', 'MB', 'PP'] #, 'PV']


#
# get command line arguments
# voter-screen.py" sd5.csv SD 5 
parser = argparse.ArgumentParser()
parser.add_argument("infile")
parser.add_argument("type")
parser.add_argument("district")
parser.add_argument("-ad", "--AD")
parser.add_argument("-sd", "--SD")
parser.add_argument("-all", "--ALL")

args = parser.parse_args()
# echo the command line
print (args.infile, args.type, args.district)

#
# function: to divide x/y and return result, protects from divide by 0
#
def div0(x,y):
    try:
        return x/y
    except ZeroDivisionError:
        return 0

#
# function: process a row of the base.csv table
#    input: row of base csv
#
def process_row(row):
     # variable definitions
    global democrats
    global republicans
    global other
    global nonps
    global iaps
    global anyvote
    global repvote
    global demvote
    global othvote
    global nonpvote
    global iapvote
    global vote

    vote = str(row['11/03/20 general'])
    voteid = str(row['StateID'])
    #print (" vote= {} {} \n".format(voteid, vote))

    if (row['Party'] == 'Democrat'):
        democrats+=1
        if vote in voted:
            anyvote+=1
            demvote+=1

    elif (row['Party'] == 'Republican'):
        republicans+=1
        if vote in voted:
            anyvote+=1
            repvote+=1
    elif (row['Party'] == 'Non-Partisan'):
        nonps+=1
        if vote in voted:
            anyvote+=1
            nonpvote+=1
    elif (row['Party'] == 'Independent American Party'):
        iaps+=1
        if vote in voted:
            anyvote+=1
            iapvote+=1
    else:
        other+=1
        if vote in voted:
            anyvote+=1
            othvote +=1
        return 0

def main():
    column = ""
    if args.type == "SD":
        column = "SenDist"
    elif args.type == "AD":
        column = "AssmDist"
    analyze_file = args.infile

    rownum = 0
    allvoters = 0
    cwd = os.getcwd()
    os.chdir(cwd)

    # read data source 
    with open(analyze_file, mode='r') as csv_file:  #'encoding='ISO-8859-1
        csv_reader = DictReader(csv_file)
        print ("opened file: {}".format(analyze_file))
        #voted = ['BR', 'EV', 'MB', 'PP', 'PV']

        # create document
        try:
            for row in csv_reader:
                if rownum == 0:
                    rownum = 1
                    continue
                rownum +=1
                if (args.type == 'ALL'):
                    pass
                elif (row[column] != args.district):
                    continue
                allvoters+=1
                #vote = str(row['11/03/20 general'])
                #print (rownum, row['CountyID'])
                
                #  process the row
                process_row(row)
                continue
        except (RuntimeError):
            print('problem occured')
            
    #  these values will be written into a table for report, for now just because
    #  VOTERS: total active voters by party group
    #   VOTED: total votes cast by party group
    # TURNOUT: turnout percentage by group

    print( 'VOTERS: {:6d} REP: {:6d} DEM: {:6d} NONP: {:6d} IAP: {:6d} OTH: {:6d}'.format(allvoters, republicans, democrats, nonps, iaps, other))
    print( ' VOTED: {:6d} REP: {:6d} DEM: {:6d} NONP: {:6d} IAP: {:6d} OTH: {:6d}'.format(anyvote, repvote, demvote, nonpvote, iapvote, othvote))
    print( '   T/O:  {:0.3f} REP:  {:0.3f} DEM:  {:0.3f} NONP:  {:0.3f} IAP:  {:0.3f} OTH:  {:0.3f}'.format(div0(anyvote,allvoters), div0(repvote,republicans), div0(demvote,democrats), div0(nonpvote,nonps), div0(iapvote,iaps), div0(othvote,other)))
    print( '   T/O:  {:0.3f} NONALIGNED:  {:0.3f}'.format(div0(anyvote,allvoters), div0(nonpvote+iapvote+othvote,nonps+iaps+other)))

    exit(0)

# --------------------------------------------------
if __name__ == '__main__':
    main()

    



    # count by party 
    
        