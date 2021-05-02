
#************************************************************************************
#                               base_precinct_split
#  purpose: 1)will split a base.csv file into many precincts 
#           2) will create a comment of the precinct statistics
#  input:   file like base.csv file
#  output:  many files containing precinct level dataFrame
#           [name of file will be precinct.nnnnnn-vvvvvv_REPnnnnn_DEMnnnnn_OTHnnnnn]                           
#                                                                                   
# *********************************************************************************** 
import numpy as np
import argparse
import errno
import csv
import os
#from csv  import reader
from csv import DictReader

#
# get command line arguments
# 
precinct_dir = ''
precinct_id = '000000'
row = ''
nextrow = ''




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
#    input: precinct file created
#   output: 2 row table with precinct statistics
#
def process_file(out_file):
     # variable definitions
     
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
    else:
        other+=1
        if vote in voted:
            anyvote+=1
            othvote +=1
        return 0
    
# 
#  create precinct directory
#    

def create_directory():
    mode = '0o771'
    global precinct_id

    os.mkdir(precinct_dir, mode)
    return 0

# 
#  create voter list file
#    

def create_voterlist():
    global row
    global nextrow

    return 0

# 
#  create summary file
#    

def create_summary():
    return 0

    
# 
#  mainline of program
#    

def main():
    rownum = 0
    global precinct_id

    file_name = ''

    parser = argparse.ArgumentParser()
    parser.add_argument("infile")
    #parser.add_argument("type")
    #parser.add_argument("district")
    #parser.add_argument("-ad", "--AD")

    args = parser.parse_args()
    # echo the command line
    print (args.infile)      # args.type, args.district)

    analyze_file = args.infile

    #    
    # reads the base.csv 
    # saves the header_row from first row
    # for each change in Precinct
    #   create a directory for that precinct to hold two files
    #   create the file of voters for that precinct 
    #       filename='precinct.nnnnnn-vvvvvv_REPnnnnn_DEMnnnnn_OTHnnnnn.csv'
    #   create a precinct summary for the precict
    #       filename='precinct.nnnnnn-vvvvvv_summary'
    #

    with open(analyze_file, mode='r', encoding='ISO-8859-1') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',', quotechar='"')
        print ("opened file: {}".format(analyze_file))

        # create new precinct 
            for next_row in csv_reader:
            if rownum == 0:
                rownum = 1
                header_row = next_row
                with open('precinct.nnnnnn-vvvvvv_REPnnnnn_DEMnnnnn_OTHnnnnn.csv', mode='w') as precinct_file:
                    precinct_writer  = csv.writer(precinct_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                    precinct_writer.writerow(header_row)
                    continue
            preinct_id = next_row[3]
            row = next_row

            if next_row[3] == precinct_id:   # test for change in precinct
                rownum +=1
                with open('precinct.nnnnnn-vvvvvv_REPnnnnn_DEMnnnnn_OTHnnnnn.csv', mode='a') as precinct_file:
                    precinct_writer  = csv.writer(precinct_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                    precinct_writer.writerow(row)
                    continue  
            # close file and start new precnct file
            else:          
                precinct_id = next_row[3]
                rownum = 0

            # create new file_name
                file_name = 'precinct.{0}-vvvvvv_REPnnnnn_DEMnnnnn_OTHnnnnn.csv'.format(precinct_id)
                with open(file_name, mode='w') as precinct_file:
                    precinct_writer  = csv.writer(precinct_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                    precinct_writer.writerow(header_row)
                
                rownum +=1
                with open(file_name, mode='a') as precinct_file:
                    precinct_writer  = csv.writer(precinct_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                    precinct_writer.writerow(row)
                    continue  




                continue


                
        process_file(precinct_file)
        #continue
        '''
        except (RuntimeError):
            print('problem occured')
        '''
            
    print( 'VOTERS: {:6d} REP: {:6d} DEM: {:6d} NONP: {:6d} IAP: {:6d} OTH: {:6d}'.format(allvoters, republicans, democrats, nonps, iaps, other))
    print( ' VOTED: {:6d} REP: {:6d} DEM: {:6d} NONP: {:6d} IAP: {:6d} OTH: {:6d}'.format(anyvote, repvote, demvote, nonpvote, iapvote, othvote))
    print( '   T/O:  {:0.3f} REP:  {:0.3f} DEM:  {:0.3f} NONP:  {:0.3f} IAP:  {:0.3f} OTH:  {:0.3f}'.format(div0(anyvote,allvoters), div0(repvote,republicans), div0(demvote,democrats), div0(nonpvote,nonps), div0(iapvote,iaps), div0(othvote,other)))
    print( '   T/O:  {:0.3f} NONALIGNED:  {:0.3f}'.format(div0(anyvote,allvoters), div0(nonpvote+iapvote+othvote,nonps+iaps+other)))

    exit(0)

# --------------------------------------------------
if __name__ == '__main__':
    main()

    