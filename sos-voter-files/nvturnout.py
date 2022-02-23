#************************************************************************************
#                                Turnout.py                                         *
#                                                                                   *
#  Input: base.csv file (preferably just after a given election)                    *
#                                                                                   *
#  Output: Analysis of turnout by party for either precinct or district             *
#                                                                                   *
# *********************************************************************************** 

from typing import Type
import pandas as pd
import numpy as np
import sys, os
import getopt
from os.path import isfile, join
import datetime
import time
import xlsxwriter
import math

BaseFile ="base.csv"

ReportFile = "turnout.xlsx"

Precinct = 0
District = 0
DType = 0
DNum = 0
Election = ""
County = ""

printFile = "print.txt"
printFileh = 0

helpReq   = 0

ProgName = "TURNOUT"                 # Name of running program
#
#=====================================================================================================================
#
#************************************
#                                   *
# Print Log line to screen and file *
#                                   *
#************************************
#
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
#
#************************************
#                                   *
#     Print command line help       *
#                                   *
#************************************
#
def printhelp():
    print ("py turnout.py  -i <basefile> -r <reportfile> -p <precinct> -d <district>")
    print ("    -i <basefile>   = Processed Secretary of state data file.")
    print ("                      Default is base.csv.")
    print ("    -r <reportfile> = output report Excel Spreadsheet file")
    print ("                      Default is turnout.xlsx")
    print ("    -c <county>     = County to report precincts from.")
    print ("                      Only needed if using multi-county base.csv file.")
    print ("    -p <precint>    = Number of a precinct to report turnout")
    print ("    -d <district>   = District to report turnout for (CDn, ADn, or SDn")
    print (" ")
    print ("               Note: Either -p or -d is required but not both.\n")
    return
#
#============================================================================================================
#
#*******************************************************
#                                                      *
#         Routine to get command line arguments        *
#                                                      *
#*******************************************************
def args(argv):
    global BaseFile, ReportFile, Precinct, District, DType, DNum, County

    print("")
    try:
        opts, args = getopt.getopt(argv,"h:i:r:p:d:c:",["help", "infile=", "reportfile=", "precinct=", "district=", "county="])
    except getopt.GetoptError:
        printhelp()
        return(2)
    for opt, arg in opts:
        if opt == '-h':
            printhelp()
            return(2)
        elif opt in ("-i", "--infile"):
            BaseFile = arg
        elif opt in ("-c", "--county"):
            County = arg.title()
        elif opt in ("-r", "--reportfile"):
            ReportFile = arg
        elif opt in ("-p", "--precinct"):
            Precinct = arg
        elif opt in ("-d", "--district"):
            District = arg
            if (len(District) < 3):
                printhelp()
                return(2)
            DType = District[0:2].upper()           # get two lettrs of district type
            DNum = District[2:]                     # get number of District
            if (DType != "AD") and (DType != "SD") and (DType !="CD"):
                printLine(f"UnKnown District Type {DType}")
                printhelp()
                return(2)
    if((Precinct != 0) and (District != 0)):
        printLine("Cannot specify both district and precinct - only one...")
        printhelp()
        return(2)
    printLine("Input base file is " + BaseFile)
    printLine("Report Spreadsheet File is " + ReportFile)
    return(0)
#
#====================================================================================================
#
#*************************************************************
#       >>>>>  M A I N   P R O G R A M   S T A R T  <<<<<    *
#*************************************************************
#
def main():
    global BaseFile, ReportFile, printFileh, Precinct, District, DType, DNum, Election, County
    #
    ###############################################################################
    #
    #  ------------ Program Starts Here --------------
    #
    StartTime = time.time()                         # get program start time
    #
    #  Open console log file (overwrite old log if it exists)
    #
    try:
        printFileh = open(printFile, "w")
    except IOError as e:
        print ("Unable to create console log file: I/O error({0}): {1}".format(e.errno, e.strerror))
        exit(2)
    except: #handle other exceptions such as attribute errors
        print ("Unexpected error:", sys.exc_info()[0])
        exit(2)
    #
    ec = args(sys.argv[1:])                         # Get command line arguments if any
    if (ec != 0):
        return (ec)                                 # get out if error
    #
    #   Next read the Older Elegible Voter file Into Dataframe
    #
    #
    #   Next read the Newer Elegible Voter file Into Dataframe
    #
    printLine(f"Loading SOS Base Data File: {BaseFile}. ")
    if (BaseFile[-4:] == ".csv"):
        base = pd.read_csv (BaseFile,low_memory=False)              #  Read .csv base data file into dataframe "base"
    else:
        base = pd.read_excel (BaseFile)                             #  Read .xls or .xlsx Eligible Voter file into dataframe "base"
    base = base.replace(np.nan, '', regex=True)                     #  make any nans into '' All data frame
    NumBase = len(base.index)                                       # get # of registered Voters in Old File
    baseheaders = list(base.columns)
    if ("County" in baseheaders):
        NoCounty = False                                            # new format, has County column
    else:
        NoCounty = True                                             # old format, no County Column
    ex = baseheaders.index("Age")+1
    Election = baseheaders[ex]
    if (NoCounty == False):
        printLine("Sorting base File on County")
        base.sort_values(by=['County'], inplace=True, kind="mergesort")
    printLine("Converting base file DataFrame to Dictionary...")
    bDict = base.to_dict(orient='list')                             # convert hframe to dictionary with parallel columns
    printLine("{0:,} Base File Voters Loaded...".format(NumBase))
    if (Precinct != 0):
        printLine(f"Calculating turnout for precinct {Precinct} for {Election} election")
    else:
        printLine(f"Calculating turnout for precincts in {DType}{DNum} for {Election} election")
    #
    #  Find Date of Election and save as datetime object
    #
    mm = Election[0:2]
    dd = Election[3:5]
    yy = Election[6:8]
    Electiondt = datetime.datetime(int("20"+yy),int(mm),int(dd))          # form election date time
    #
    #  Define Data Accumulator variables
    #
    Pcts = []                   # list of precincts
    #
    RepActive = []               # By precinct list of # of Republicans who could have voted in this election
    RepVoted = []                # By precinct list of # of Republicans who did vote in this election
    #
    DemActive = []
    DemVoted = []
    #
    NPPActive = []
    NPPVoted = []
    #
    IAPActive = []
    IAPVoted = []
    #
    OTHActive = []
    OTHVoted= []
    #
    if (Precinct != 0):
        #
        #  Doing a single precint, init lists with a single entry each
        #
        Pcts.append(int(Precinct))
        RepActive.append(0)
        RepVoted.append(0)
        DemActive.append(0)
        DemVoted.append(0)
        NPPActive.append(0)
        NPPVoted.append(0)
        IAPActive.append(0)
        IAPVoted.append(0)
        OTHActive.append(0)
        OTHVoted.append(0)
    #
    #  Now index through base file.  For each entry where base["Status"] == "Active"
    # 
    #   1.  If registration date is newer than election date, ignore the record
    #   2.  Count as Active Registration that could have voted
    #   3.  If they voted, Count as having actually voted in this election
    #
    bx = 0                  # index into Old Voter Dictionary
    while (bx < NumBase):
        if ((bx % 2000) == 0):
            print (f"Processing record {bx}\r", end="")
        Pct = int(base["Precinct"][bx])
        if (Pct == 0):
            bx += 1
            continue                                            # ignore date record in base.csv
        if (base["Status"][bx] == "Inactive"):
            bx += 1
            continue                                            # ignore inactive voter records
        if (Precinct != 0):
            if (Pct != int(Precinct)):
                bx += 1
                continue                                        # this voter not in selected precinct
            else:
                lx = 0                                          # this voter selected, list index = 0
        else:
            if (DType == "SD"):
                dist = str(base["SenDist"][bx])
                dsave = dist
            elif (DType == "AD"):
                dist = str(base["AssmDist"][bx])
                dsave = dist
            elif (DType == "CD"):
                dist = str(base["CongDist"][bx])
                dsave = dist
            if(dist == ""):
                bx = bx+1
                continue                                        # ignore voter with no district given
            if (len(dist) > 2):
                if (dist[-2] == "."):
                    dist = dist[0:-2]                           # Handle number ending in .0
                else:
                    dist = dist[2:]                             # strip district type letters if present
            #
            #  Dist is now district number as a string
            #  see if district we want
            #
            try:
                if (int(dist) != int(DNum)):
                    bx = bx+1
                    continue                                        # voter not in selected district
            except:
                #
                #  Caught another district format that we haven't handled yet
                #
                temp = base["StateID"][bx]
                print(f"Dist = \"{dist}\", StateID={temp}, record = {bx+2}")
                exit(2)
            #
            #  This voter is in the selected district, see if in selected county (or 1st county in base.csv)
            #
            if (NoCounty == False):
                if (County == ""):
                    County = base["County"][bx].title()                # set county if not specified
                else:
                    if (County != base["County"][bx].title()):
                        bx = bx+1
                        continue                                        # voter not in selected county - ignore
            #
            #  See if registered in time to vote in this election
            #
            rdate = base["RegDate"][bx]
            if (rdate[1] == "/"):
                rdate = "0"+rdate
            if(rdate[4] == "/"):
                rdate = rdate[0:3] + "0" + rdate[3:]
            mm = rdate[0:2]
            dd = rdate[3:5]
            yy = rdate[6:10]
            try:
                rdate = datetime.datetime(int(yy),int(mm),int(dd))          # form election date time
            except:
                rdate = base["RegDate"][bx]
                print(f"Date Error: input = {rdate}\n mm = {mm}, dd = {dd}, yyyy = {yy}" )
                exit(2)
            if (rdate > Electiondt):
                bx=bx+1
                continue                                                    # registered after election, ignore
            #
            #  Set lx = to list index of this precinct
            #       if this is first voter in this precinct
            #       add item to accumulator lists and set lx to added item
            #
            try:
                lx = Pcts.index(Pct)                            # set index if precinct in list
            except:
                #
                #  Doing a new precint in this district, add entry to lists
                #
                Pcts.append(Pct)                                # add this Precinct
                RepActive.append(0)                             # with no voters so far
                RepVoted.append(0)
                DemActive.append(0)
                DemVoted.append(0)
                NPPActive.append(0)
                NPPVoted.append(0)
                IAPActive.append(0)
                IAPVoted.append(0)
                OTHActive.append(0)
                OTHVoted.append(0)
                lx = len(Pcts)-1                                # set index to added entries
        #
        #  lx is now index into lists for this voter record
        #
        party = base["Party"][bx]
        votecode = base[Election][bx]
        if (party == "Republican"):
            RepActive[lx] += 1
            if (votecode != ""):
                RepVoted[lx] += 1
        elif (party == "Democrat"):
            DemActive[lx] += 1
            if (votecode != ""):
                DemVoted[lx] += 1
        elif (party == "Independent American Party"):
            IAPActive[lx] += 1
            if (votecode != ""):
                IAPVoted[lx] += 1
        elif (party == "Non-Partisan"):
            NPPActive[lx] += 1
            if (votecode != ""):
                NPPVoted[lx] += 1
        else:
            OTHActive[lx] += 1
            if (votecode != ""):
                OTHVoted[lx] += 1
        bx=bx+1                                             # bump to next voter
    #
    #  Calculations done, now report results
    #
    if (len(Pcts) == 0):
        printLine(f">>>>  No selected precinsts found in {BaseFile} !!!")
        return(2)
    #
    #  have data to report, make the spreadsheet
    #
    printLine(f"Creating Excel Spreadsheet {ReportFile} ...")
    #
    # Create report workbook and add a worksheet.
    #
    workbook = xlsxwriter.Workbook(ReportFile)
    worksheet = workbook.add_worksheet()
    #
    # set those workbook properties that we can at the start
    #
    worksheet.set_landscape()                                           # set to print in landscape orientation
    worksheet.set_paper(1)                                              # 5 for legal paper, 1 for Letter Paper
    worksheet.set_margins(left=0.7, right=0.7, top=0.75, bottom=0.75)   # set margins for printing
    worksheet.fit_to_pages(1, 0)                                        # print 1 page wide and as long as necessary.
    # set column widths
    worksheet.set_column('A:P', 11.71)                                  # All Columns to 11.71 units wide
    #
    #  Create the cell formats we will need for this spreadsheet
    #
    fmt_left_bold = workbook.add_format({'bold': True , 'align': 'left', 'font_size': '13'})
    fmt_center_bold = workbook.add_format({'bold': True , 'align': 'center', 'font_size': '13'})
    fmt_right_bold = workbook.add_format({'bold': True , 'align': 'right', 'font_size': '13'})
    fmt_number = workbook.add_format({'bold': False , 'align': 'right', 'font_size': '13','num_format': '#,##0'})
    fmt_number_pink = workbook.add_format({'bold': False , 'align': 'right', 'bg_color': '#F8CBAD', 'font_size': '13', 'num_format': '#,##0'})
    fmt_pct = workbook.add_format({'bold': False , 'align': 'right', 'font_size': '13','num_format': '0.0%'})
    fmt_right = workbook.add_format({'bold': False , 'align': 'right', 'font_size': '13'})
    #
    #
    #  Basic spreadsheet structure is now set, and the formats we need are defined.
    #  Next write all of the fixed text cells with proper formatting including
    #  merging any cells we need to merge.
    #
    #----------------------------------------------------------------------------
    #   Titles for Columns and Rows
    #
    if (Precinct != 0):
        temp = "Turnout for " + County + " County Precinct " + str(Precinct) + " for the " + Election + " Election"
    else:
        temp = "Turnout for " + County + " County Precincts in " + DType + str(DNum) + " for the " + Election +" Election"
    worksheet.merge_range('A1:P1', temp , fmt_center_bold)
    worksheet.write('A3', 'Precinct', fmt_center_bold)                                        # Row 3 is Column Title Text
    worksheet.write('B3', 'REP Reg', fmt_center_bold)
    worksheet.write('C3', 'REP Voted', fmt_center_bold)
    worksheet.write('D3', 'REP T/O %', fmt_center_bold)
    worksheet.write('E3', 'DEM Reg', fmt_center_bold)
    worksheet.write('F3', 'DEM Voted', fmt_center_bold)
    worksheet.write('G3', 'DEM T/O %', fmt_center_bold)
    worksheet.write('H3', 'NPP Reg', fmt_center_bold)
    worksheet.write('I3', 'NPP Voted', fmt_center_bold)
    worksheet.write('J3', 'NPP T/O %', fmt_center_bold)
    worksheet.write('K3', 'IAP Reg', fmt_center_bold)
    worksheet.write('L3', 'IAP Voted', fmt_center_bold)
    worksheet.write('M3', 'IAP T/O %', fmt_center_bold)
    worksheet.write('N3', 'OTH Reg', fmt_center_bold)
    worksheet.write('O3', 'OTH Voted', fmt_center_bold)
    worksheet.write('P3', 'OTH T/O %', fmt_center_bold)
    #
    order = np.argsort(Pcts)                                            # get list of sorted precinct index order to report parallel lists in
    j=4                                                                 # 1st spreadsheet row number for precincts
    for i in order:                                                     # fetch list items in sorted order
        row=str(j)                                                      # get spreadsheet row # as string
        temp = Pcts[i]
        if (temp > 19999):                                              # check for 6 digit washoe precinct numbers
            temp = int(temp/100)                                        # convert Washoe precincts with two trailing 0's to 4 digits
        worksheet.write('A'+row, temp, fmt_right_bold)
        worksheet.write('B'+row, RepActive[i], fmt_number)
        worksheet.write('C'+row, RepVoted[i], fmt_number)
        if (RepActive[i] != 0):
            worksheet.write('D'+row, RepVoted[i]/RepActive[i], fmt_pct)
        worksheet.write('E'+row, DemActive[i], fmt_number)
        worksheet.write('F'+row, DemVoted[i], fmt_number)
        if (DemActive[i] != 0):
            worksheet.write('G'+row, DemVoted[i]/DemActive[i], fmt_pct)
        worksheet.write('H'+row, NPPActive[i], fmt_number)
        worksheet.write('I'+row, NPPVoted[i], fmt_number)
        if (NPPActive[i] != 0):
            worksheet.write('J'+row, NPPVoted[i]/NPPActive[i], fmt_pct)
        worksheet.write('K'+row, IAPActive[i], fmt_number)
        worksheet.write('L'+row, IAPVoted[i], fmt_number)
        if (IAPActive[i] != 0):
            worksheet.write('M'+row, IAPVoted[i]/IAPActive[i], fmt_pct)
        worksheet.write('N'+row, OTHActive[i], fmt_number)
        worksheet.write('O'+row, OTHVoted[i], fmt_number)
        if (OTHActive[i] != 0):
            worksheet.write('P'+row, OTHVoted[i]/OTHActive[i], fmt_pct)
        j += 1                                                          # Bump Row Number
    #
    #  Now Write Totals Row formulas
    #
    lrow= str(j-1)                                                                  # last precinct row number
    trow = str(j)                                                                   # Get Row Number for totals as character string
    worksheet.write('A'+trow, 'Totals', fmt_right_bold)         # Row Title
    worksheet.write_formula('B'+trow, '=SUM(B4:B' + lrow + ')', fmt_number)         # total REP Active Column
    worksheet.write_formula('C'+trow, '=SUM(C4:C' + lrow + ')', fmt_number)         # total REP Voted Column
    worksheet.write_formula('D'+trow, '=C' + trow + '/B' + trow, fmt_pct)           # total REP T/O %
    worksheet.write_formula('E'+trow, '=SUM(E4:E' + lrow + ')', fmt_number)         # total DEM Active Column
    worksheet.write_formula('F'+trow, '=SUM(F4:F' + lrow + ')', fmt_number)         # total DEM Voted Column
    worksheet.write_formula('G'+trow, '=F' + trow + '/E' + trow, fmt_pct)           # total DEM T/O %
    worksheet.write_formula('H'+trow, '=SUM(H4:H' + lrow + ')', fmt_number)         # total NPP Active Column
    worksheet.write_formula('I'+trow, '=SUM(I4:I' + lrow + ')', fmt_number)         # total NPP Voted Column
    worksheet.write_formula('J'+trow, '=I' + trow + '/H' + trow, fmt_pct)           # total NPP T/O %
    worksheet.write_formula('K'+trow, '=SUM(K4:K' + lrow + ')', fmt_number)         # total IAP Active Column
    worksheet.write_formula('L'+trow, '=SUM(L4:L' + lrow + ')', fmt_number)         # total IAP Voted Column
    worksheet.write_formula('M'+trow, '=L' + trow + '/K' + trow, fmt_pct)           # total IAP T/O %
    worksheet.write_formula('N'+trow, '=SUM(N4:N' + lrow + ')', fmt_number)         # total OTH Active Column
    worksheet.write_formula('O'+trow, '=SUM(O4:O' + lrow + ')', fmt_number)         # total OTH Voted Column
    worksheet.write_formula('P'+trow, '=O' + trow + '/N' + trow , fmt_pct)          # total OTH T/O %
    #
    #  Now close and write .xlsx file
    #
    try:
        workbook.close()
    except Exception as e:
        print('>>>>> Error Writing Report Spreadsheet file!!\n   Message, {m}'.format(m = str(e)))
    EndTime = time.time()
    print (f"Done! - Total Elapsed time is {int((EndTime - StartTime)*10)/10} seconds.\n")
    return(0)

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
    exit (main())