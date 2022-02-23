
import pandas as pd
import numpy as np
import sys, getopt, os
import logging


base = "base-la.csv"

#*******************************************************
#                                                      *
#  Routine to get command line arguments (if any)      *
#                                                      *
#*******************************************************



#**********************************************
#    M A I N   P R O G R A M   S T A R T      *
#**********************************************
#
def main():
    global base
    
    baserows = 0
    
    logging.basicConfig(format='%(asctime)s %(levelname)-8s %(message)s',
            level=logging.INFO,
            datefmt='%Y-%m-%d %H:%M:%S')

    #
    #  Can expand this to check file name to see if .csv or .xls and make each
    #  read use either read_csv or read_excel as needed to allow full
    #  flexibility in input files.
    #
    #  For now, output file is always a .csv file.
    #
    base = pd.read_csv (base,low_memory=False)      #  Read SOS base.csv file into DataFrame "base"
    baserows=len(base.index)
    #
    # get lists of columnn lbels from teh three input files
    #
    basehead=list(base.columns)                        # get SOS data column labels
    outhead = basehead  
    
    dict = ''
    inc = -1
                                           # get precinct extract file column labels
    

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