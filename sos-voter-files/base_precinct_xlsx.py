#************************************************************************************
#                          base_precinct_xlsx.py                                    *
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
            "LikelyToVote",
            "LastVote",
            "Score"
]
base=""                                     # SOS data Dataframe Object (loaded at start of pgm)
SnglPct=0                                   # Extract single Precinct option if non-null
subdir=""                                   # output subdirectory

#*******************************************************
#                                                      *
#  Routine to get command line arguments (if any)      *
#                                                      *
#*******************************************************
def printhelp():
    print('base_precinct_xlsx.py -s <Sosfile> -p <precint> -d <outdir>')
    print('     -s <Sosfile>  = Secretary of State base file (from nvvoter.py)')
    print('                     default is base.csv in current directory.')
    print('     -p <precinct> = single precinct to extract.')
    print('                     default is all precincts in base file.')
    print('     -d <outdir>   = directory to put extracted files into.')
    print('                     default is current working directory.')
    print
    return(0)
#
#
def args(argv):
    global Sosfile, SnglPct, subdir
    print("")
    try:
        opts, args = getopt.getopt(argv,"h:s:p:d:",["help", "sosfile=", "precinct=", "--dir"])
    except getopt.GetoptError:
        printhelp()
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            printhelp()
            sys.exit()
        elif opt in ("-s", "--sosfile"):
            Sosfile = arg
        elif opt in ("-d", "--directory"):
            subdir=arg
        elif opt in ("-p", "--precinct"):
            SnglPct = arg
            print(f"Extracting single precinct {SnglPct}")
    print("Input SOS data file is " + Sosfile)

#**********************************************
#    M A I N   P R O G R A M   S T A R T      *
#**********************************************
#
def main():
    global base, SnglPct,  subdir, outfile
    StartTime = time.time()                         # get start time
    args(sys.argv[1:])                              # Get command line arguments if any
    Dir = os.getcwd()                               # Start output directory with our current working directory
    if (subdir != ""):
        if(subdir[1] != ":"):
           Dir = os.path.join(Dir,subdir)           # if a Survey Subdirectory specified add it to current path
        else:
            Dir = subdir                            # fully qualified directory specified, use it as specified
    outfile = os.path.join(Dir,outfile)             # form full temp file name
    print (f"Extracting precinct files to Directory = {Dir}")
    if(os.path.isdir(Dir) == False):
        os.mkdir(Dir)                               # create diretory if it doesn't exist
    #
    #  Can expand this to check file name to see if .csv or .xls and make each
    #  read use either read_csv or read_excel as needed to allow full
    #  flexibility in input files.
    #
    #  For now, output file is always a .csv file.
    #
    print(f"Loading {Sosfile} ... ", end="", flush=True)
    base = pd.read_csv (Sosfile,low_memory=False)   #  Read SOS base.csv file into DataFrame "base"
    EndTime = time.time()
    print (f"{Sosfile} load took {int((EndTime - StartTime)*10)/10} seconds\n")
    baserows=len(base.index)
    #
    # get lists of column labels from the input file
    # and calculate indexes for needed columns
    #
    basehead=list(base.columns)                     # get SOS data column labels
    hits = 0                                        # say no column title found yet
    x = 0                                           # index to 1st column
    for item in basehead:
        if (item == "CountyID"):
            CountyId = x
            hits += 1
        if (item == "First"):
            First = x
            hits += 1
        if (item == "Middle"):
            Middle = x
            hits += 1
        if (item == "Last"):
            Last = x
            hits += 1
        if (item == "Phone"):
            PhoneNum = x
            hits += 1
        if (item == "RegDate"):
            RegDate = x
            hits += 1
        if (item == "Party"):
            RegParty = x
            hits += 1
        if (item == "StreetNo"):
            StreetNo = x
            hits += 1
        if (item == "StreetName"):
            Streetname = x
            hits += 1
        if (item == "RegisteredDays"):
            RegDays = x
            hits += 1
        if (item == "Age"):
            VAge = x
            Cycle = VAge+1                       # offset to 1st election cycle
            Elections = x+1
            hits += 1
        if (item == "LikelytoVote"):
            LikelytoVote = x
            hits += 1
        if (item == "Score"):
            Score = x
            hits += 1
        x += 1
    if (hits != 13):
        print (f"SOS data file format error - only found {hits} required columns")
        print(basehead)
        exit (2)
    #
    #  Create a list of precinct #s in pctlist
    #
    pctlist = []
    if (SnglPct != 0):
        pctlist.append(int(SnglPct))                        # extracting only 1 precinct
    else:
        for item in base["Precinct"]:                       # build list of all precincts in base.csv to extract
            if (item == 0):
                continue                                    # skip special date of SOS data record
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

        # set workbook print properties
        worksheet.set_landscape()                                           # set to print in landscape orientation
        worksheet.set_paper(5)                                              # 5 for legal paper  (1 for Letter Paper)
        worksheet.set_margins(left=0.7, right=0.7, top=0.75, bottom=0.75)   # set print margins to Excel "normal"
        worksheet.fit_to_pages(1, 0)                                        # print 1 page wide and as long as necessary.
        worksheet.repeat_rows(0)                                            # Repeat the header row on each printed page.
        worksheet.set_footer('&CPage &P of &N')                             # set printed page footer

        # set column widths that are always the same
        worksheet.set_column(0, 0, 8.43)         # CountyID
        worksheet.set_column(4, 4, 14.71)        # Phone
        worksheet.set_column(5, 5, 10)           # RegDate
        worksheet.set_column(6, 6, 26.29)        # Party
        worksheet.set_column(7, 7, 8.43)         # StreetNo
        worksheet.set_column(9, 9, 7.71)         # RegDays
        worksheet.set_column(10, 10, 5.43)       # Age
        worksheet.set_column(11, 11, 11.86)      # LikelyToVote
        worksheet.set_column(12, 12, 10)         # LastVote
        worksheet.set_column(13, 13, 5.43)       # Score
        #
        #  Columns 2, 3, 4, and 9 width set at end once we know longest entry in each
        #  Set max length for each to 0 here to start out.
        MaxFirst = 0                             # col 2 initial width
        MaxLast = 0                              # col 3 initial width
        MaxMiddle = 0                            # col 4 initial width
        MaxStreet = 0                            # col 9 initial width
        #
        #  Create cell formats we will need to use for this spreadsheet
        #
        fmt_left = workbook.add_format({'bold': False , 'align': 'left'})
        fmt_center = workbook.add_format({'bold': False , 'align': 'center'})
        fmt_right = workbook.add_format({'bold': False , 'align': 'right'})
        # Add a bold format to use for header cells.
        header_bold = workbook.add_format({'bold': True , 'align': 'center'})
        # Add a date format for cells with dates
        date_format = workbook.add_format({'num_format': 'mm/dd/yy;@', 'align': 'center'})

        #  Write out formatted header row for this precinct .xlsx file
        x=0
        for item in outheader:
            worksheet.write(0, x, item , header_bold)
            x += 1

        # Now write out extracted base.csv dataframe rows for this precinct.
        #    1. Pick only the base.csv columns we want for extract.xlsx.
        #    2. Format cells as they are written.
        #    3. Keep track of longest text entry in First, Last, Middle and StreetName
        #       columns to allow setting column widths before finshing output.
        #    4. Keep count of Rep, Dem and Other voters to allow creating
        #       file names for each precinct with those counts in name.
        #    5. Close extract.xlsx and rename to formatted precinct file name
        row=0
        for x in range(count):
            outrow = list(extract.iloc[x])                              # get next row of dataframe as list
            i=0
            for item in outrow:
                if (item != item):                                      # Make any NAN cells null strings
                    outrow[i] = ""
                i += 1
            party = outrow[RegParty]                                    # Fetch Party of this voter
            if (party == "Democrat"):
                numDem=numDem + 1                                       # Add to number of Democrat Voters
                continue                                                # don't write out record
            row = row+1                                                 # write to next row
            worksheet.write_number (row, 0, outrow[CountyId], fmt_right)  # CountyID
            worksheet.write_string (row, 1, outrow[First], fmt_left)    # First
            L = len(outrow[First])
            if (L > MaxFirst):
                MaxFirst = L                                            # longest so far
            worksheet.write_string (row, 2, outrow[Last], fmt_left)     # Last
            L = len(outrow[Last])
            if (L > MaxLast):
                MaxLast = L                                             # longest so far
            worksheet.write_string (row, 3, outrow[Middle], fmt_left)   # Middle
            L = len(outrow[Middle])
            if (L > MaxMiddle):
                MaxMiddle = L                                           # longest so far
            phone=outrow[PhoneNum]
            if (phone.isdigit()):
                worksheet.write_number (row, 4, int(phone), fmt_right)  # Phone is numeric
            elif (phone == ""):
                worksheet.write_blank(row, 4, None)                     # No Phone Number
            else:
                worksheet.write (row, 4, phone, fmt_right)              # Phone not numeric
            worksheet.write (row, 5, outrow[RegDate], date_format)      # Regdate
            worksheet.write_string (row, 6, party, fmt_left)            # Party
            if (party == "Republican"):
                numRep=numRep + 1                                       # Add to # Republican Voters
            else:
                numOth = numOth + 1                                     # Add to number of "Other Party" Voters
            snum = outrow[StreetNo]
            if isinstance(snum,str):
                worksheet.write_blank(row, 7, None)                     # Street Number is string
            elif snum == "":
                worksheet.write_blank(row, 7, None)                     # Street Number is blank
            else:
                worksheet.write_number (row, 7, int(snum), fmt_right)   # Street Number is numeric
            SName = outrow[Streetname]
            if (SName == ""):
                worksheet.write_blank(row, 8, None)                     # Street Name Blank
            else:
                worksheet.write (row, 8, outrow[Streetname], fmt_left)  # Street Name
            L = len(outrow[Streetname])
            if (L > MaxStreet):
                MaxStreet = L                                           # longest so far
            Days = outrow[RegDays]
            if (Days == ""):
                print("RegDays Blank in Row " + str(row+1) + " Precinct " + str(PctNum))
                worksheet.write_blank(row, 9, None)
            else:
                worksheet.write_number (row, 9, outrow[RegDays], fmt_right)     # Reg Days
            Age = outrow[VAge]
            if (Age == ""):
                print("Age Blank in Row " + str(row+1) + " Precinct " + str(PctNum))
                worksheet.write_blank(row, 10, None)
            else:
                worksheet.write_number (row, 10, Age, fmt_right)        # Age
            worksheet.write_string (row, 11, outrow[LikelytoVote], fmt_left)      # Likely To Vote
            LastVote = "Never"
            for i in range(20):
                if (outrow[Elections+i] != ""):
                    LastVote = basehead[Cycle+i]
                    break
            if (LastVote != "Never"):
                LastVote = LastVote[0:6] + "20" + LastVote[6:8]             # truncate to date only and expand year to 4 digits
            worksheet.write_string (row, 12, LastVote, date_format)     # latest Election Voted In
            worksheet.write_number (row, 13, outrow[Score], fmt_right)  # Score
        #
        # We've built precinct spreadsheet in memory now do final work and write it out
        #         
        # set First, Last, Middle and Street Name column widths
        worksheet.set_column(1, 1, MaxFirst*.96)                        # "First" column width
        worksheet.set_column(2, 2, MaxLast*.96)                         # "Last" column width
        worksheet.set_column(3, 4, MaxMiddle*.96)                       # "Middle" column width
        worksheet.set_column(8, 8, MaxStreet*.96)                       # "StreetName" column width
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
        outname = os.path.join(Dir , qualname)
        os.replace(outfile,outname)                                     # rename extract.xlsx to actual file name
    #    exit()                                              # >>>>>>   Debug exit after 1 precinct  <<<<           
    print("\nPrecinct .xlsx File(s) Extracted.")
    EndTime = time.time()
    TotSec = int((EndTime - StartTime)*10)/10
    TotMin = int (TotSec/60)
    if (TotMin > 0):
        TotSec = int((TotSec -(TotMin*60))*10)/10
        print(f"Total Elapsed time is {TotMin} Minutes {TotSec} seconds.\n")
    else:
        print(f"Total Elapsed time is {TotSec} seconds.\n")
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
