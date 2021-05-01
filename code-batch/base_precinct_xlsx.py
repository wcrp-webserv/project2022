#************************************************************************************
#                          base_precinct_xlsx.py                                    
#                                                                                   *
#  Input is Processed Secretary of State base.csv file.                             *
#                                                                                   *
#  Output is an series of extracted files, one for each precinct in base.csv.       *
#                                                                                   *
#  Output file naming is PCT_nn_TOTtt_REPrr_DEMdd_OTHoo.CSV                         *
#        nn = precint number                                                        *
#        tt = total # of voters                                                     *
#        rr = # of Republican voters                                                *
#        dd = # of Democrat voters                                                  *
#        oo = # voters of other or no party                                         *
# *********************************************************************************** 

import pandas as pd
import numpy as np
import sys, getopt, os
import xlsxwriter
from datetime import datetime
import time

Sosfile = "base.csv"                       # Secretary of State Data with voting results combined
outfile = "extract.xlsx"                   # output extended member file
outheader = ["CountyID",
            "Precinct",
            "First",
            "Last",
            "Middle",
            "Phone",
            "RegDate",
            "Party",
            "StreetNo",
            "StreetName",
            "RegDays",
            "Age",
            "LikelyToVote"
]
base=""                                     # SOS data Dataframe Object (loaded at start of pgm)
SnglPct=0                                   # Extract single Precinct option if non-null

#*******************************************************
#                                                      *
#  Routine to get command line arguments (if any)      *
#                                                      *
#*******************************************************
def args(argv):
    global Sosfile, SnglPct
    try:
        opts, args = getopt.getopt(argv,"h:s:p:",["help", "sosfile=", "precinct="])
    except getopt.GetoptError:
        print('base_precinct_xlsx.py -s <Sosfile> -p <precint>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print('base_precinct_xlsx.py -s <Sosfile> -p <precint#>')
            sys.exit()
        elif opt in ("-s", "--sosfile"):
            Sosfile = arg
        elif opt in ("-p", "--precinct"):
            SnglPct = arg
            if (len(SnglPct) < 5):
                SnglPct = SnglPct + "00"
            print(f"Extracting single precinct {SnglPct}")
    print("Input files:")
    temp = '   SOS data file is "' + Sosfile + '"'
    print(temp)
    print("")

#**********************************************
#    M A I N   P R O G R A M   S T A R T      *
#**********************************************
#
def main():
    global base, SnglPct
    StartTime = time.time()                         # get start time
    args(sys.argv[1:])                              #  Get command line arguments if any
    #
    #  Can expand this to check file name to see if .csv or .xls and make each
    #  read use either read_csv or read_excel as needed to allow full
    #  flexibility in input files.
    #
    #  For now, output file is always a .csv file.
    #
    base = pd.read_csv (Sosfile,low_memory=False)   #  Read SOS base.csv file into DataFrame "base"
    EndTime = time.time()
    print (f"Runtime to load .csv file {EndTime - StartTime}")
    baserows=len(base.index)
    #
    # get lists of columnn labels from the input file
    #
    basehead=list(base.columns)                     # get SOS data column labels
     #
    #  Create a list of precinct #s in pctlist
    #
    pctlist = []
    if (SnglPct != 0):
        pctlist.append(int(SnglPct))                        # extracting only 1 precinct
    else:
        for item in base["Precinct"]:                       # build list of all precincts in base.csv to extract
            if item not in pctlist:
                pctlist.append(item)                        # add precinct to list
        pctlist.sort()                                      # sort list in ascending order
        print (f"Found {len(pctlist)} Precincts to Extract")
    #
    # For each precinct, create an extraction of base.csv items only for that precinct
    #
    for PctNum in pctlist:
        extract = base.loc[base["Precinct"] == PctNum]                      # extract dataframe of entries for this dataframe
        count = len(extract.index)                                          # count = How many entries in extracted dataframe
        print ("Precinct " + str(PctNum) + " has " + str(count) + " Rows")  # Print what we found for this precinct
        #
        #  extract is now a dataframe in the same format as base.csv but containing
        #  only those rows that have the column "Precinct" matching PctNum
        #
        #  open precinct output file with temp name extract.csv
        #
        numRep=0                                                            # init counters for this precinct
        numDem=0
        numOth=0
        row=0                                                               # start at row 0
        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook(outfile)
        worksheet = workbook.add_worksheet()

        # set workbook properties that we can at the start
        worksheet.set_landscape()                                           # set to print in landscape orientation
        worksheet.set_paper(5)                                              # 5 for legal paper, 1 for Letter Paper
        worksheet.set_margins(left=0.7, right=0.7, top=0.75, bottom=0.75)   # set margins for printing

        # set column widths that are always the same
        worksheet.set_column(0, 0, 8.43)         # CountyID
        worksheet.set_column(1, 1, 7.43)         # Precinct
        worksheet.set_column(5, 5, 14.71)        # Phone
        worksheet.set_column(6, 6, 10)           # RegDate
        worksheet.set_column(7, 7, 26.29)        # Party
        worksheet.set_column(8, 8, 8.43)         # StreetNo
        worksheet.set_column(10, 10, 7.71)       # RegDays
        worksheet.set_column(11, 11, 5.43)       # Age
        worksheet.set_column(12, 12, 11.86)      # LikelyToVote
        #
        #  Columns 2, 3, 4, and 9 set at end once we know longest entry in each
        #  Set max length for each to 0 here to start out
        MaxFirst = 0
        MaxLast = 0
        MaxMiddle = 0
        MaxStreet = 0

        #  Create cell formats we will need to use for this spreadsheet
        fmt_left = workbook.add_format({'bold': False , 'align': 'left'})
        fmt_center = workbook.add_format({'bold': False , 'align': 'center'})
        fmt_right = workbook.add_format({'bold': False , 'align': 'right'})
        # Add a bold format to use for header cells.
        header_bold = workbook.add_format({'bold': True , 'align': 'center'})
        # Add a date format for cells with dates
        date_format = workbook.add_format({'num_format': 'mm/dd/yy;@', 'align': 'center'})

        #  Write out formatted header row for extract.xlsx
        x=0
        for item in outheader:
            worksheet.write(0, x, item , header_bold)
            x += 1
        worksheet.repeat_rows(0)                    # Repeat the header row on each printed page.
        worksheet.set_footer('&CPage &P of &N')     # set printed page footer

        # Now write out extracted base.csv dataframe rows for this precinct.
        #    1. Pick only the base.csv columns we want for extract.xlsx.
        #    2. Format cells as they are written.
        #    3. Keep track of longest text entry in First, Last, Middle and StreetName
        #       columns to allow setting column widths before finshing output.
        #    4. Keep count of Rep, Dem and Other voters to allow creating
        #       file names for each precinct with those counts in name.
        #    5. Close extract.xlsx and rename to formatted precinct file name
        for x in range(count):
            outrow = list(extract.iloc[x])                              # get next row of dataframe as list
            i=0
            for item in outrow:
                if (item != item):                                      # Make any NAN cells null strings
                    outrow[i] = ""
                i += 1
            row = x+1
            worksheet.write_number (row, 0, outrow[0], fmt_right)       # CountyID
            worksheet.write_number (row, 1, outrow[3], fmt_right)       # Precinct
            worksheet.write_string (row, 2, outrow[7], fmt_left)        # First
            L = len(outrow[7])
            if (L > MaxFirst):
                MaxFirst = L                                            # longest so far
            worksheet.write_string (row, 3, outrow[8], fmt_left)        # Last
            L = len(outrow[8])
            if (L > MaxLast):
                MaxLast = L                                             # longest so far
            worksheet.write_string (row, 4, outrow[9], fmt_left)        # Middle
            L = len(outrow[9])
            if (L > MaxMiddle):
                MaxMiddle = L                                           # longest so far
            phone=outrow[11]
            if (phone.isdigit()):
                worksheet.write_number (row, 5, int(phone), fmt_right)  # Phone is numeric
            elif (phone == ""):
                worksheet.write_blank(row, 5, None)                     # No Phone Number
            else:
                worksheet.write (row, 5, phone, fmt_right)              # Phone not numeric
            worksheet.write (row, 6, outrow[14], date_format)           # Regdate
            party = outrow[15]                                          # Fetch Party of this voter
            worksheet.write_string (row, 7, party, fmt_left)            # Party
            if (party == "Republican"):
                numRep=numRep + 1                                       # Add to # Republican Voters
            elif (party == "Democrat"):
                numDem=numDem + 1                                       # Add to number of Democrat Voters
            else:
                numOth = numOth + 1                                     # Add to number of "Other Party" Voters
            snum = outrow[16]
            if (snum == ""):
                worksheet.write_blank(row, 8, None)                     # Street Number is blank
            else:
                worksheet.write_number (row, 8, int(snum), fmt_right)   # Street Number is numeric
            SName = outrow[17]
            if (SName == ""):
                worksheet.write_blank(row, 9, None)                     # Street Name Blank
            else:
                worksheet.write (row, 9, outrow[17], fmt_left)          # Street Name
            L = len(outrow[17])
            if (L > MaxStreet):
                MaxStreet = L                                           # longest so far
            Days = outrow[24]
            if (Days == ""):
                print("RegDays Blank in Row " + str(row+1) + " Precinct " + str(PctNum))
                worksheet.write_blank(row, 10, None)
            else:
                worksheet.write_number (row, 10, outrow[24], fmt_right)     # Reg Days
            Age = outrow[25]
            if (Age == ""):
                print("Age Blank in Row " + str(row+1) + " Precinct " + str(PctNum))
                worksheet.write_blank(row, 11, None)
            else:
                worksheet.write_number (row, 11, Age, fmt_right)        # Age
            worksheet.write_string (row, 12, outrow[53], fmt_left)      # Likely To Vote
        #
        # We've built precinct spreadsheet in memory now do final work and write it out
        #         
        # set First, Last, Middle and Street Name column widths
        worksheet.set_column(2, 2, MaxFirst*.96)                        # "First" column width
        worksheet.set_column(3, 3, MaxLast*.96)                         # "Last" column width
        worksheet.set_column(4, 4, MaxMiddle*.96)                       # "Middle" column width
        worksheet.set_column(9, 9, MaxStreet*.96)                       # "StreetName" column width
        #
        #  Now build the correct file name for this precinct file
        # 
        qualname = "PCTID_"+ str(PctNum) + "_TOT_" + str(count) + "_REP" + str(numRep)
        qualname = qualname + "_DEM" + str(numDem) + "_OTH" + str(numOth) + ".xlsx"
        #
        worksheet.set_header('&C' + qualname)                           # set filename as print header
        #
        workbook.close()                                                # Write out and close extract.xlsx temp spreadsheet
        #
        os.replace(outfile,qualname)                                     # rename extract.xlsx to actual file name
#        exit()                                              # >>>>>>   Debug exit after 1 precinct  <<<<           
    print("Precinct .xlsx Files Extracted...exiting")
    EndTime = time.time()
    print (f"Total Elapsed time is {EndTime - StartTime}")
    exit(0)


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
