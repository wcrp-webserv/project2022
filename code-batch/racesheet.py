#************************************************************************************
#                                   racesheet.py                                    *
#                                                                                   *
#  Input is Processed Secretary of State base.csv file.                             *
#                                                                                   *
#  Output is an excel spreadsheet detailing the assessment of the race.             *
#                                                                                   *
#  Output file naming is RaceDDNN.xlsx                                              *
#        DD = District Type selected to analyze (SD, AD, ...)                       *
#        NN = Number of district being Analyzed                                     *
#                                                                                   *
#             example output file name: raceAD31.xlsx                               *
#                                                                                   *
# *********************************************************************************** 

import pandas as pd
import numpy as np
import sys, getopt, os
import xlsxwriter
from datetime import datetime, date
import time

Sosfile = "base.csv"                        # Secretary of State Data with voting results combined
outfile = "raceSD31.xlsx"                   # output extended member file
base=""                                     # SOS data Dataframe Object (loaded at start of pgm)
SOSdate = "??/??/????"                      # "as-of" date of Secretary of State date base.csv compiled from
CompDate = "??/??/????"                     # date base.csv was compiled from SOS data
Dist=""                                     # district to analyze (ADnn or SDnn)
Candidate = ""                              # Candidate we are doing analysis for (default = "Candidate" + "Dist")
cycle = ""                                  # election cycle to analyze
histdir = ""                                # History Data File directory
bDict=[]                                    # Dictionary for base.csv column access by text title (built when base.csv opened)
Cfgfile = "RaceCfg.xlsx"                    # Party Voting Percentages Configuration Spreadsheet

#*******************************************************
#                                                      *
#  Routine to get command line arguments (if any)      *
#                                                      *
#*******************************************************
def printhelp():
    print('racesheet.py -d <district> -c <CandidateName> -y <electionyear> -s <Sosfile> -h <histdir>')
    print('    <district> = AD or SD district to do race sheet for - REQUIRED')
    print('    <CandidateName> = Name of candidate sheet generated for:')
    print('        default is generic \"CandidateDist\" if none specified')
    print('    <electionyear> = Election Year Sheet is for.')
    print('        default is next election cycle')
    print('    <Sosfile> = Secretary of State data for at least district in base.csv format')
    print('        default is file base.csv in current working directory')
    print('    <histdir> = directory containing required election history data')
    print('        default is current working directory')
    return(0)

def args(argv):
    global Sosfile, outfile, Dist, Candidate, cycle, histdir
    print("")
    try:
        opts, args = getopt.getopt(argv,"?:s:d:c:y:h:",["help", "sosfile=", "district=", "candidate=", "year=", "history="])
    except getopt.GetoptError:
        printhelp()
        exit(2)
    for opt, arg in opts:
        if opt == '-?':
            printhelp()
            sys.exit()
        elif opt in ("-s", "--sosfile"):
            Sosfile = arg
        elif opt in ("-y", "--year"):
            cycle = arg
        elif opt in ("-c", "--candidate"):
            Candidate = arg
        elif opt in ("-h", "--history"):
            histdir = arg
        elif opt in ("-d", "--district"):
            Dist = arg
            Dist= Dist.upper()
            outfile = "race" + Dist + ".xlsx"
    if (Dist == ""):
        print (">>>> No District Selected, Nothing to do...\n")
        printhelp()
        exit(2)
    if (Candidate == ""):
        Candidate = "Candidate" + Dist                          # no candidate specified, default to Generic Candidate
    if (cycle == ""):
        todays_date = date.today()                              # No Cycle Specified, Default to next election cycle
        cycle = todays_date.year                                # Get Current Year
        cycle = int(cycle)                                      # Default to next election cycle before cmd line args fetched
        if ((cycle % 2) != 0):
            cycle = cycle + 1                                   # elections in even years only
        cycle = str(cycle)                                      # convert Default Cycle Year to string
    return(0)

#
#  ROutine to set Name/Vote nan entries to "" or 0
#
def FixVoteAndName(Candidates,Votes):
    for x in range(1, len(Candidates)):                     #Empty Candidates = ""
        if str(Candidates[x]) == 'nan':
            Candidates[x] = ""
        else:
            temp = Candidates[x]                            # Truncate to last name only
            pos = temp.index(',')                           # Find comma in string
            temp = temp[:pos]                               # substrting out the last name
            Candidates[x] = temp                            # replace with last name only

    for x in range(1, len(Votes)):                          # Empty Votes = 0
        if str(Votes[x]) == 'nan':
            Votes[x] = 0
    return

#**********************************************
#    M A I N   P R O G R A M   S T A R T      *
#**********************************************
#
def main():
    global base, outfile, Sosfile, Dist, Candidate, cycle, histdir, bDict
    global SOSdate, CompDate

    StartTime = time.time()                                 # get start time (to calc program run time)
    #
    args(sys.argv[1:])                                      # Get command line arguments if any
    #
    # Verify District is SD or AD
    #
    DType=Dist[:2]                                          # get AD or SD
    DNum = int(Dist[2:])                                    # Get District Number
    if (DType != "SD"):
        if (DType != "AD"):
            print (f">>> District Must Be AD or SD, You Gave {DType} - Aborting...")
            exit(2)
        else:
            if ((DNum < 1) or (DNum > 42)):
                print(f">>> Assembly districts go from 1 to 42, you gave {Dist} - Aborting...")
                exit(2)
    else:
        if ((DNum < 1) or (DNum > 21)):
            print(f">>> Senate districts go from 1 to 21, you gave {Dist} - Aborting...")
            exit(2)
    #
    Dir = os.getcwd()                                       # Start output directory with our current working directory
    outfile = os.path.join(Dir,outfile)                     # form full spreadsheet file name
    if (histdir != ""):
        if(histdir[1] != ":"):
            histdir = os.path.join(Dir,histdir)             # if a History File Subdirectory specified add it to current path
    else:
        histdir = Dir                                       # History must be in current working directory
    print(f"histdir: {histdir}, Dir: {Dir}")
    #
    print("Doing Analysis for district " + Dist + ", Candidate " + Candidate + " for the " + cycle + " election cycle.")
    print("Output report file is " + outfile)
    print(f"Loading Configuration file {Cfgfile}")
    try:
        cdf = pd.read_excel (Cfgfile)                       # read configuration into DataFram cdf
    except Exception as e:
        print('>>> Error Opening file {0}!!\n>>> Message, {1}'.format(Cfgfile, str(e)))
        exit(2)
    #
    # Find % Conservative and %T/O estimates in configuration file
    #
    cdfhead = list(cdf.columns)                             # get configuration header
    ToCfg = ""
    ConsCfg = ""
    for i, row in cdf.iterrows():
        if "conservative" in row["Party"].lower():
            ConsCfg = row
        if ("t/o" in row["Party"].lower()):
            ToCfg = row
    if (type(ToCfg) == type("")):
        print("Configuration File Missing Required Row %T/O - aborting...")
        exit(2)
    if (type(ConsCfg) == type("")):
        print("Configuration File Missing Required Row %Conservative - aborting...")
        exit(2)
    #
    #  Open the Secretary of State base.csv file
    #
    print(f"Loading SOSfile {Sosfile} ...")
    try:
        base = pd.read_csv (Sosfile,low_memory=False)           #  Read SOS base.csv file into DataFrame "base"
    except Exception as e:
        print('>>> Error Opening file {0}!!\n>>> Message, {1}'.format(Sosfile, str(e)))
        exit(2)
    basehead=list(base.columns)                                 # get SOS data column labels
    bDict = {basehead[i]: i for i in range(len(basehead))}  # dictionary of indexes to names for base.csv row
    #
    #  Build date this race sheet was created (today)
    #
    dt = datetime.today()                                      # get current date as datetime object
    CompDate = str(dt.month) + "/" + str(dt.day) + "/" + str(dt.year)   # turn mm/dd/yyyy string
    #
    #  If base.csv has special first row, get SOS as-of date and base.csv compiled on date
    #
    if ((base.iloc[0]["CountyID"] == 0) and (base.iloc[0]["Precinct"] == 0)):
        SOSdate = base.iloc[0]["RegDate"]                               # get SOS "as-of" date as mm/dd/yyy string
    #
    #  Extract district race being analyzed from base.csv into dataframe Dbase
    #
    print("Extracting records for selected district...")
    if (DType == 'SD'):
        Dbase = base.loc[base["SenDist"] == DNum]           # extract entries for requested SD
    else:
        Dbase = base.loc[base["AssmDist"] == DNum]          # extract entires for requested AD
    del base                                                # release base.csv dataframe memory
    #
    #  now working with dataframe of only base.csv records for district being analyzed
    #
    baserows=len(Dbase.index)                               # Number of rows in extracted dataframe
    print ("{0:,} Registered Voters for District {1}{2}".format(baserows, DType, DNum))
    #
    #   Dbase is now a base.csv format dataframe with registered voters for this district
    #
    Prev1 = int(cycle)-2                                    # calculate previous four election cycle years
    Prev2 = int(cycle)-4
    Prev3 = int(cycle)-6
    Prev4 = int(cycle)-8
    if ((int(cycle) % 4) == 0):
        Presidential = 1                                    # analyzing a Presidential Year Election
    else:
        Presidential = 0                                    # analyzing an Off Year Election
    HistPrev1 = str(Prev1) + "-" + Dist[:2] + "Votes.xlsx"  # Prev1 Election History File
    HistPrev1= os.path.join(histdir,HistPrev1)
    HistPrev2 = str(Prev2) + "-" + Dist[:2] + "Votes.xlsx"  # Prev2 Election History File
    HistPrev2= os.path.join(histdir,HistPrev2)
    HistPrev3 = str(Prev3) + "-" + Dist[:2] + "Votes.xlsx"  # Prev3 Election History File
    HistPrev3= os.path.join(histdir,HistPrev3)
    HistPrev4 = str(Prev4) + "-" + Dist[:2] + "Votes.xlsx"  # Prev4 Election History File
    HistPrev4= os.path.join(histdir,HistPrev4)
    HistTO = os.path.join(histdir,"2010-2020 RegHistoryByParty.xlsx") # Registration for AD, SD and County by election History File
    #
    # Load 1 cycle back history
    #
    print( f"Loading {Prev1} History File: {HistPrev1}")
    try:
        dfPrev1 = pd.read_excel(HistPrev1)                  # Load History file for 1st past election
        extract = dfPrev1.loc[dfPrev1[DType] == Dist]       # Extract entries for requested District
        if extract.empty:
            print('>>>>> No vote for {0}{1} in {2}, Leaving This Cycle Out.'.format(DType, DNum, Prev1))
            P1Candidates = [""] * 10
            P1Votes = [0] * 10
            P1Votes[4] = 8
        else:
            P1Candidates = list(extract.iloc[1])                # candidates list
            P1Votes = list(extract.iloc[3])                     # Vote Totals
            for x in range(2, len(P1Candidates)):               #Empty Candidates = ""
                if str(P1Candidates[x]) == 'nan':
                    P1Candidates[x] = ""
                else:
                    temp = P1Candidates[x]                      # there is a name, fetch it
                    if (len(temp) > 1):
                        pos = temp.index(',')                   # find the first comma (end of last name)
                        P1Candidates[x] = temp[:pos]            # extract the last name and put only that back
            for x in range(2, len(P1Votes)):                    # Empty Votes = 0
                if str(P1Votes[x]) == 'nan':
                    P1Votes[x] = 0
    except Exception as e:
        print('>>>>> Error Opening file!!\n   Message, {m}'.format(m = str(e)))
        print('>>>>> Leaving This Cycle Out.')
        P1Candidates = [""] * 10
        P1Votes = [0] * 10
        P1Votes[4] = 8
    #
    # Load 2 Cycle Back History
    #
    print( f"Loading {Prev2} History File: {HistPrev2}")
    try:
        dfPrev2 = pd.read_excel(HistPrev2)
        extract = dfPrev2.loc[dfPrev2[DType] == Dist]       # Extract entries for requested District
        if extract.empty:
            print('>>>>> No vote for {0}{1} in {2}, Leaving This Cycle Out.'.format(DType, DNum, Prev2))
            P2Candidates = [""] * 10
            P2Votes = [0] * 10
            P2Votes[4] = 8
        else:
            P2Candidates = list(extract.iloc[1])                # candidates list
            P2Votes = list(extract.iloc[3])                     # Vote Totals
            for x in range(2, len(P2Candidates)):               #Empty Candidates = ""
                if str(P2Candidates[x]) == 'nan':
                    P2Candidates[x] = ""
                else:
                    temp = P2Candidates[x]                      # there is a name, fetch it
                    if (len(temp) > 1):
                        pos = temp.index(',')                   # find the first comma (end of last name)
                        P2Candidates[x] = temp[:pos]            # extract the last name and put only that back
            for x in range(2, len(P2Votes)):                    # Empty Votes = 0
                if str(P2Votes[x]) == 'nan':
                    P2Votes[x] = 0
    except Exception as e:
        print('>>>>> Error Opening file!!\n   Message, {m}'.format(m = str(e)))
        print('>>>>> Leaving This Cycle Out.')
        P2Candidates = [""] * 10
        P2Votes = [0] * 10
        P2Votes[4] = 8
    #
    # Load 3 Cycle Back History
    #
    print( f"Loading {Prev3} History File: {HistPrev3}")
    try:
        dfPrev3 = pd.read_excel(HistPrev3)
        extract = dfPrev3.loc[dfPrev3[DType] == Dist]           # Extract entries for requested District
        if extract.empty:
            print('>>>>> No vote for {0}{1} in {2}, Leaving This Cycle Out.'.format(DType, DNum, Prev3))
            P3Candidates = [""] * 10
            P3Votes = [0] * 10
            P3Votes[4] = 8
        else:
            P3Candidates = list(extract.iloc[1])                    # candidates list
            P3Votes = list(extract.iloc[3])                         # Vote Totals
            for x in range(2, len(P3Candidates)):                   #Empty Candidates = ""
                if str(P3Candidates[x]) == 'nan':
                    P3Candidates[x] = ""
                else:
                    temp = P3Candidates[x]                          # there is a name, fetch it
                    if (len(temp) > 1):
                        pos = temp.index(',')                       # find the first comma (end of last name)
                        P3Candidates[x] = temp[:pos]                # extract the last name and put only that back
            for x in range(2, len(P3Votes)):                        # Empty Votes = 0
                if str(P3Votes[x]) == 'nan':
                    P3Votes[x] = 0
    except Exception as e:
        print('>>>>> Error Opening file!!\n   Message, {m}'.format(m = str(e)))
        print('>>>>> Leaving This Cycle Out.')
        P3Candidates = [""] * 10
        P3Votes = [0] * 10
        P3Votes[4] = 8
    #
    # Load 4 Cycle Back History
    #
    print( f"Loading {Prev4} History File: {HistPrev4}")
    try:
        dfPrev4 = pd.read_excel(HistPrev4)
        extract = dfPrev4.loc[dfPrev4[DType] == Dist]           # Extract entries for requested District
        if extract.empty:
            print('>>>>> No vote for {0}{1} in {2}, Leaving This Cycle Out.'.format(DType, DNum, Prev4))
            P4Candidates = [""] * 10
            P4Votes = [0] * 10
            P4Votes[4] = 8
        else:
            P4Candidates = list(extract.iloc[1])                    # candidates list
            P4Votes = list(extract.iloc[3])                         # Vote Totals
            for x in range(2, len(P4Candidates)):                   #Empty Candidates = ""
                if str(P4Candidates[x]) == 'nan':
                    P4Candidates[x] = ""
                else:
                    temp = P4Candidates[x]                          # there is a name, fetch it
                    if (len(temp) > 1):
                        pos = temp.index(',')                       # find the first comma (end of last name)
                        P4Candidates[x] = temp[:pos]                # extract the last name and put only that back
            for x in range(2, len(P4Votes)):                        # Empty Votes = 0
                if str(P4Votes[x]) == 'nan':
                    P4Votes[x] = 0
    except Exception as e:
        print('>>>>> Error Opening file!!\n   Message, {m}'.format(m = str(e)))
        print('>>>>> Leaving This Cycle Out.')
        P4Candidates = [""] * 10
        P4Votes = [0] * 10
        P4Votes[4] = 8
    #
    # Load District Registered Voter History
    #
    print( f"Loading Registered Voter By Party History File: {HistTO}")
    try:
        dfTO = pd.read_excel(HistTO)
    except Exception as e:
        print('>>>>> Error Opening file!!\n   Message, {m}'.format(m = str(e)))
        exit(2)
    print ("Election History Files Loaded, Building Spreadsheet ...")
    #
    #  Can expand this to check file name to see if .csv or .xls and make each
    #  read use either read_csv or read_excel as needed to allow full
    #  flexibility in input files.
    #
    #  For now, output file is always a .csv file.
    #
    #print(f"Loading {Sosfile} ... ", end="", flush=True)
    #base = pd.read_csv (Sosfile,low_memory=False)   #  Read SOS base.csv file into DataFrame "base"
    #EndTime = time.time()
    #print (f"{Sosfile} load took {int((EndTime - StartTime)*10)/10} seconds\n")
    #baserows=len(base.index)
    #
    # get lists of column labels from the input file
    #
    #basehead=list(base.columns)                     # get SOS data column labels
     #
    # For each precinct, create an extraction of base.csv items only for that precinct
    #
    # Create report workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(outfile)
    worksheet = workbook.add_worksheet()

    # set workbook properties that we can at the start
    worksheet.set_landscape()                                           # set to print in landscape orientation
    worksheet.set_paper(1)                                              # 5 for legal paper, 1 for Letter Paper
    worksheet.set_margins(left=0.7, right=0.7, top=0.75, bottom=0.75)   # set margins for printing
    worksheet.fit_to_pages(1, 0)                                        # print 1 page wide and as long as necessary.
    worksheet.print_area('A1:Q45')                                      # Set Print Area ( what cells to print)
    worksheet.set_footer('&CPage &P of &N')                             # set printed page footer
    worksheet.set_zoom(75)                                              # set Excel viewing zoom at 75%

    # set column widths
    worksheet.set_column('A:A', 17.57)          # 1st column wider
    worksheet.set_column('B:Q', 10.38)          # rest of columns all 10.38 wide
    #
    #  Create all of the different cell formats we will need for this spreadsheet
    #
    fmt_left_bold = workbook.add_format({'bold': True , 'align': 'left', 'font_size': '14'})
    fmt_left_gray = workbook.add_format({'bold': True , 'align': 'left', 'bg_color': '#D9D9D9', 'font_size': '14'})
    fmt_center_bold = workbook.add_format({'bold': True , 'align': 'center', 'font_size': '14'})
    fmt_center_gray = workbook.add_format({'bold': True , 'align': 'center', 'bg_color': '#D9D9D9', 'font_size': '14'})
    fmt_number = workbook.add_format({'bold': False , 'align': 'right', 'font_size': '14','num_format': '#,##0'})
    fmt_number_pink = workbook.add_format({'bold': False , 'align': 'right', 'bg_color': '#F8CBAD', 'font_size': '14', 'num_format': '#,##0', 'border' : 1})
    fmt_pct = workbook.add_format({'bold': False , 'align': 'right', 'font_size': '14','num_format': '0.0%'})
    fmt_pct_center_pink = workbook.add_format({'bold': False , 'align': 'center', 'font_size': '14','num_format': '0.0%', 'bg_color': '#F8CBAD', 'border' : 1})
    fmt_pct_center_yel = workbook.add_format({'bold': False , 'align': 'center', 'font_size': '14','num_format': '0.0%', 'bg_color': '#FFFF00', 'border' : 1})
    fmt_number_center_org = workbook.add_format({'bold': False , 'align': 'center', 'bg_color': '#FFC000', 'font_size': '14', 'num_format': '#,##0', 'border' : 1})
    fmt_number_center_grn = workbook.add_format({'bold': False , 'align': 'center', 'bg_color': '#92D050', 'font_size': '14', 'num_format': '#,##0', 'border' : 1})
    fmt_center_border = workbook.add_format({'bold': False , 'align': 'center', 'border' : 1, 'font_size': '14'})
    fmt_center_bold_border = workbook.add_format({'bold': True , 'align': 'center', 'border' : 1, 'font_size': '14'})
    fmt_number_border = workbook.add_format({'bold': False , 'align': 'right', 'font_size': '14','num_format': '#,##0', 'border' : 1})
    fmt_pct_border = workbook.add_format({'bold': False , 'align': 'right', 'font_size': '14','num_format': '0.0%', 'border' : 1})
    fmt_center_gray_border = workbook.add_format({'bold': True , 'align': 'center', 'bg_color': '#D9D9D9', 'font_size': '14', 'border' : 1})
    fmt_center_border = workbook.add_format({'bold': False , 'align': 'center', 'top' : 1, 'border' : 1, 'font_size': '14'})
    fmt_center_border2L = workbook.add_format({'bold': False , 'align': 'center', 'top' : 1, 'bottom' : 1, 'left' : 5, 'right' : 1, 'font_size': '14'})
    fmt_center_border2R = workbook.add_format({'bold': False , 'align': 'center', 'top' : 1, 'bottom' : 1, 'left' : 1, 'right' : 5, 'font_size': '14'})
    fmt_center_bold_border2 = workbook.add_format({'bold': True , 'align': 'center', 'top' : 1, 'bottom' : 1, 'left' : 5, 'right' : 5, 'font_size': '14'})
    fmt_center_bold_border2L = workbook.add_format({'bold': True , 'align': 'center', 'top' : 1, 'bottom' : 1, 'left' : 5, 'right' : 1, 'font_size': '14'})
    fmt_center_bold_border2R = workbook.add_format({'bold': True , 'align': 'center', 'top' : 1, 'bottom' : 1, 'left' : 1, 'right' : 5, 'font_size': '14'})
    fmt_number_center_border = workbook.add_format({'bold': False , 'align': 'center', 'border' : 1, 'font_size': '14','num_format': '#,##0'})
    fmt_number_border2L = workbook.add_format({'bold': False , 'align': 'right', 'font_size': '14','num_format': '#,##0',
                                              'top' : 1, 'bottom' : 1, 'left' : 5, 'right' : 1})
    fmt_number_border2R = workbook.add_format({'bold': False , 'align': 'right', 'font_size': '14','num_format': '#,##0',
                                              'top' : 1, 'bottom' : 1, 'left' : 1, 'right' : 5})
    fmt_number_center_border2 = workbook.add_format({'bold': False , 'align': 'center', 'font_size': '14','num_format': '#,##0',
                                                    'top' : 1, 'bottom' : 1, 'left' : 5, 'right' : 5})
    #
    #
    #  Basic spreadsheet structure is now set, and the formats we need are defined.
    #  Next write all of the fixed text cells with proper formatting including
    #  merging any cells we need to merge.
    #
    #----------------------------------------------------------------------------
    #   Fixed values come from a RaceCfg.xlsx file
    #
    # Row 3 is Projected T/O % by party
    # Row 4 is Projected Conservative % by party
    # Row 5 is Projected Progressive % by party
    # 
    #
    try:
        worksheet.write('B3', ToCfg["REP"], fmt_pct)                      # Percent Projected T/O Factor REP
        worksheet.write('C3', ToCfg["NP"], fmt_pct)
        worksheet.write('D3', ToCfg["IAP"], fmt_pct)
        worksheet.write('E3', ToCfg["LP"], fmt_pct)
        worksheet.write('F3', ToCfg["GP"], fmt_pct)
        worksheet.write('G3', ToCfg["OTH"], fmt_pct)
        worksheet.write('H3', ToCfg["DEM"], fmt_pct)
        worksheet.write('B4', ConsCfg["REP"], fmt_pct)                      # Percent Projected Conservative Vote REP
        worksheet.write('C4', ConsCfg["NP"], fmt_pct)
        worksheet.write('D4', ConsCfg["IAP"], fmt_pct)
        worksheet.write('E4', ConsCfg["LP"], fmt_pct)
        worksheet.write('F4', ConsCfg["GP"], fmt_pct)
        worksheet.write('G4', ConsCfg["OTH"], fmt_pct)
        worksheet.write('H4', ConsCfg["DEM"], fmt_pct)
    except Exception as e:
        print('>>>>> Missing Configuration File Column Title {m}'.format(m = str(e)))
    worksheet.write_formula('B5', '=1.0-B4', fmt_pct)                       # Progressive Vote Factors = 100% - Conservative %
    worksheet.write_formula('C5', '=1.0-C4', fmt_pct)
    worksheet.write_formula('D5', '=1.0-D4', fmt_pct)
    worksheet.write_formula('E5', '=1.0-E4', fmt_pct)
    worksheet.write_formula('F5', '=1.0-F4', fmt_pct)
    worksheet.write_formula('G5', '=1.0-G4', fmt_pct)
    worksheet.write_formula('H5', '=1.0-H4', fmt_pct)
    #
    #----------------------------------------------------------------------------
    #   Titles for Assessment by Registration Section
    #
    temp = Dist + " " + Candidate                                   # form 1st header row text
    worksheet.merge_range('A1:H1', temp , fmt_center_gray)
    temp = "Created: " + CompDate
    worksheet.merge_range('I1:Q1', temp , fmt_center_gray)
    temp = cycle + " Assessment by Registration"                    # form 2nd header row text
    worksheet.merge_range('A2:H2', temp , fmt_center_gray)
    temp = "Registration Data As Of: " + SOSdate
    worksheet.merge_range('I2:Q2', temp , fmt_center_gray)
    worksheet.write('A3', '%T/O  est.', fmt_left_bold)               # Row text labels
    worksheet.write('A4', '%Cons est.', fmt_left_bold)
    worksheet.write('A5', '%Prog est.', fmt_left_bold)
    worksheet.merge_range('L5:O5', 'Resulting Projected Election', fmt_center_bold_border)
    #
    worksheet.write('A6', 'Party', fmt_left_bold)                   # Row 6 is Titles
    worksheet.write('B6', 'REP', fmt_center_bold)
    worksheet.write('C6', 'NP', fmt_center_bold)
    worksheet.write('D6', 'IAP', fmt_center_bold)
    worksheet.write('E6', 'LP', fmt_center_bold)
    worksheet.write('F6', 'GP', fmt_center_bold)
    worksheet.write('G6', 'OTH', fmt_center_bold)
    worksheet.write('H6', 'DEM', fmt_center_bold)
    worksheet.write('I6', 'TOTAL', fmt_center_bold)
    worksheet.write('J6', 'AllOther', fmt_center_bold)
    worksheet.write('L6', 'R-D', fmt_center_bold_border)
    worksheet.write('M6', 'CONS', fmt_center_bold_border)
    worksheet.write('N6', 'PROG', fmt_center_bold_border)
    worksheet.write('O6', 'C-P', fmt_center_bold_border)
    #
    worksheet.write('A7', 'Reg Voters', fmt_left_bold)              # Row 7 = Registered Voters
    worksheet.write('A8', 'Act Voters', fmt_left_bold)              # Row 8 = Active Voters
    worksheet.write('A9', 'Active %', fmt_left_bold)                # Row 9 = Active %
    worksheet.write('L9', ' ', fmt_center_border)                   # two blank bordered cells on right
    worksheet.write('O9', ' ', fmt_center_border)
    worksheet.write('A10','Proj Votes', fmt_left_bold)              # Row 10 = Projected Votes
    #
    #----------------------------------------------------------------------------
    #   Titles for Assessment of Prior Election Section
    #
    worksheet.merge_range('A12:Q12', 'Assessment of Prior Elections' , fmt_center_gray)     # Row 12 = Section Title
    worksheet.write('A13', 'Election Year', fmt_left_bold)
    worksheet.merge_range('B13:E13', Prev1, fmt_center_bold_border2) # Row 13 = Prior Election year titles
    worksheet.merge_range('F13:I13', Prev2, fmt_center_bold_border2)
    worksheet.merge_range('J13:M13', Prev3, fmt_center_bold_border2)
    worksheet.merge_range('N13:Q13', Prev4, fmt_center_bold_border2)
    #
    worksheet.write('A14', 'Party', fmt_left_bold)                  # Row 14 = Labels for Assessment of Prior Election Section
    worksheet.write('B14', 'REP', fmt_center_bold_border2L)                       # Complete Parties Header Line
    worksheet.write('C14', 'R-D', fmt_center_bold_border)
    worksheet.write('D14', 'DEM', fmt_center_bold_border)
    worksheet.write('E14', 'OTH', fmt_center_bold_border2R)
    worksheet.write('F14', 'REP', fmt_center_bold_border2L)
    worksheet.write('G14', 'R-D', fmt_center_bold_border)
    worksheet.write('H14', 'DEM', fmt_center_bold_border)
    worksheet.write('I14', 'OTH', fmt_center_bold_border2R)
    worksheet.write('J14', 'REP', fmt_center_bold_border2L)
    worksheet.write('K14', 'R-D', fmt_center_bold_border)
    worksheet.write('L14', 'DEM', fmt_center_bold_border)
    worksheet.write('M14', 'OTH', fmt_center_bold_border2R)
    worksheet.write('N14', 'REP', fmt_center_bold_border2L)
    worksheet.write('O14', 'R-D', fmt_center_bold_border)
    worksheet.write('P14', 'DEM', fmt_center_bold_border)
    worksheet.write('Q14', 'OTH', fmt_center_bold_border2R)
    worksheet.write('A15', 'Candidates', fmt_left_bold)
    worksheet.write('A16', 'Votes', fmt_left_bold)
    worksheet.write('A17', 'UnderVote', fmt_left_bold)
    worksheet.write('A18', 'Total Votes', fmt_left_bold)
    #
    #----------------------------------------------------------------------------
    #   Titles for Assessment of Prior Year(s) Turnout Section
    #
    worksheet.merge_range('A20:Q20', 'Prior Year(s) Turnout' , fmt_center_gray) # Row 20 = Section Header Row
    worksheet.write('A21', 'Election Year', fmt_left_bold)
    worksheet.merge_range('B21:D21', Prev1, fmt_center_bold_border)             # Row 21 = Prior Election year titles
    worksheet.merge_range('F21:H21', Prev2, fmt_center_bold_border)
    worksheet.merge_range('J21:L21', Prev3, fmt_center_bold_border)
    worksheet.merge_range('N21:P21', Prev4, fmt_center_bold_border)
    worksheet.write('A22', 'Party', fmt_left_bold)                              # Row 22 = Titles Labels for Prior Year T/O Section
    worksheet.write('B22', 'REP', fmt_center_bold_border)
    worksheet.write('C22', 'OTH', fmt_center_bold_border)
    worksheet.write('D22', 'DEM', fmt_center_bold_border)
    worksheet.write('F22', 'REP', fmt_center_bold_border)
    worksheet.write('G22', 'OTH', fmt_center_bold_border)
    worksheet.write('H22', 'DEM', fmt_center_bold_border)
    worksheet.write('J22', 'REP', fmt_center_bold_border)
    worksheet.write('K22', 'OTH', fmt_center_bold_border)
    worksheet.write('L22', 'DEM', fmt_center_bold_border)
    worksheet.write('N22', 'REP', fmt_center_bold_border)
    worksheet.write('O22', 'OTH', fmt_center_bold_border)
    worksheet.write('P22', 'DEM', fmt_center_bold_border)
    worksheet.write('A23', 'Act Voters', fmt_left_bold)                         # Row 23 = Active Voters by party
    worksheet.write('A24', 'Tot Act Voters', fmt_left_bold)                     # Row 24 = Total Active Voters
    worksheet.write('A25', '%Turnout', fmt_left_bold)                           # Row 25 = %Turnout
    worksheet.write('B27', '4 CYC', fmt_center_gray_border)                     # Row 27 = Calculated T/O headers
    worksheet.write('C27', '2 Like CYC', fmt_center_gray_border)
    worksheet.write('A28', 'Calc Avg T/O', fmt_left_bold)                       # Row 28 = Calculated Average T/O
    #
    #----------------------------------------------------------------------------
    #   Titles for Current Election Win Number Section
    #
    temp = cycle + " Win Number Calculation"                                    # form section header text
    worksheet.merge_range('A30:Q30', temp, fmt_center_gray)                     # Row 30 = Section Header
    worksheet.merge_range('F31:H31', 'Last 4 Cycles', fmt_center_gray_border)   # Row 31 = Cycle Box Headers
    worksheet.merge_range('K31:M31', 'Previous 2 Like Cycles', fmt_center_gray_border)
    worksheet.merge_range('A32:E32', 'Average Turnout (est. from prior year T/O', fmt_left_bold)
    worksheet.merge_range('A33:E33', 'Total Active Voters (from registration analysis)', fmt_left_bold)
    worksheet.merge_range('A34:E34', 'Expected Turnout (Active Voters * Avg TO)', fmt_left_bold)
    worksheet.merge_range('A35:E35', 'Win Number (votes over half to win)', fmt_left_gray)
    worksheet.merge_range('F35:H35', 'plus 100', fmt_center_gray_border)        # complete Win Number fixed text
    worksheet.merge_range('K35:M35', 'plus 100', fmt_center_gray_border)
    worksheet.merge_range('A36:E36', 'For 2 person race, Votes needed to Win', fmt_left_bold)
    #
    #----------------------------------------------------------------------------
    #   Titles for Current Election Simulation Section
    #
    temp = cycle + " Election Simulation"                               # Row 38 = 1st Header Row
    worksheet.merge_range('A38:E38', temp, fmt_center_gray)
    temp = cycle + " with last 4-cycles"
    worksheet.merge_range('F38:H38', temp, fmt_center_gray_border)
    worksheet.merge_range('I38:J38', ' ', fmt_center_gray)
    temp = cycle + " with like 2-cycles"
    worksheet.merge_range('K38:M38', temp, fmt_center_gray_border)
    worksheet.merge_range('N38:Q38', ' ', fmt_center_gray)
    worksheet.write('F39', 'REP', fmt_center_gray_border)               # Row 39 = Simulation Party Headers
    worksheet.write('G39', 'OTH', fmt_center_gray_border)
    worksheet.write('H39', 'DEM', fmt_center_gray_border)
    worksheet.write('K39', 'REP', fmt_center_gray_border)
    worksheet.write('L39', 'OTH', fmt_center_gray_border)
    worksheet.write('M39', 'DEM', fmt_center_gray_border)
    worksheet.merge_range('A39:E39', 'Party', fmt_left_bold)
    worksheet.merge_range('A40:E40', 'Active Voters', fmt_left_bold)    # Row 40 = Active Voters
    worksheet.merge_range('A41:E41', 'Turnout', fmt_left_bold)          # Row 41 = Turnout
    worksheet.merge_range('A42:E42', 'Projected OTH Split', fmt_left_bold)  # Row 42 = Split of Other Parties
    worksheet.merge_range('A43:E43', 'Likely Outcome', fmt_left_bold)   # Row 43 = Likely Outcome
    worksheet.write('A44', 'Mod/Wk REP', fmt_left_bold)                 # Row 44 = Mod/Wk Republicans
    worksheet.write('A45', 'Strong OTH', fmt_left_bold)                 # Row 45 = Strong Other Party Voters
    #
    #---------------------------------------------------------------------
    #
    #   Now fill in the formula cells
    #
    worksheet.write_formula('I7', '=SUM(B7:H7)', fmt_number)            # Registered Voter Totals
    worksheet.write_formula('J7', '=SUM(C7:G7)', fmt_number)
    worksheet.write_formula('L7', '=B7-H7', fmt_number_border)                 # Registered R-D
    worksheet.write_formula('M7', '=B4*B7+C4*C7+D4*D7+E4*E7+F4*F7+G4*G7+H4*H7', fmt_number_border)
    worksheet.write_formula('N7', '=B5*B7+C5*C7+D5*D7+E5*E7+F7*F5+G5*G7+H5*H7', fmt_number_border)
    worksheet.write_formula('O7', '=M7-N7', fmt_number_border)
    #
    worksheet.write_formula('I8', '=SUM(B8:H8)', fmt_number)            # Active Voter Totals
    worksheet.write_formula('J8', '=SUM(C8:G8)', fmt_number)
    worksheet.write_formula('L8', '=B8-H8', fmt_number_border)                 # Active Voter R-D  
    worksheet.write_formula('M8', '=B4*B8+C4*C8+D4*D8+E4*E8+F4*F8+G4*G8+H4*H8', fmt_number_border)
    worksheet.write_formula('N8', '=B5*B8+C5*C8+D5*D8+E5*E8+F8*F5+G5*G8+H5*H8', fmt_number_border)
    worksheet.write_formula('O8', '=M8-N8', fmt_number_border)
    #
    worksheet.write_formula('B9', '=B8/B7', fmt_pct)                    # active % by party
    worksheet.write_formula('C9', '=C8/C7', fmt_pct)
    worksheet.write_formula('D9', '=D8/D7', fmt_pct)
    worksheet.write_formula('E9', '=E8/E7', fmt_pct)
    worksheet.write_formula('F9', '=F8/F7', fmt_pct)
    worksheet.write_formula('G9', '=G8/G7', fmt_pct)
    worksheet.write_formula('H9', '=H8/H7', fmt_pct)
    worksheet.write_formula('I9', '=I8/I7', fmt_pct)
    worksheet.write_formula('J9', '=J8/J7', fmt_pct)
    worksheet.write_formula('M9', '=M8/M7', fmt_pct_border)
    worksheet.write_formula('N9', '=N8/N7', fmt_pct_border)
    #
    worksheet.write_formula('B10', '=B8*B3', fmt_number)                # Projected Votes by party
    worksheet.write_formula('C10', '=C8*C3', fmt_number)
    worksheet.write_formula('D10', '=D8*D3', fmt_number)
    worksheet.write_formula('E10', '=E8*E3', fmt_number)
    worksheet.write_formula('F10', '=F8*F3', fmt_number)
    worksheet.write_formula('G10', '=G8*G3', fmt_number)
    worksheet.write_formula('H10', '=H8*H3', fmt_number)
    worksheet.write_formula('I10', '=SUM(B10:H10)', fmt_number)
    worksheet.write_formula('L10', '=B10-H10', fmt_number_border)              # Projected R-D
    worksheet.write_formula('M10', '=B4*B10+C4*C10+D4*D10+E4*E10+F4*F10+G4*G10+H4*H10', fmt_number_border)
    worksheet.write_formula('N10', '=B5*B10+C5*C10+D5*D10+E5*E10+F10*F5+G5*G10+H5*H10', fmt_number_border)
    worksheet.write_formula('O10', '=M10-N10', fmt_number_pink)
    #
    worksheet.write_formula('C16', '=B16-D16', fmt_number_pink)         # prior election R-D
    worksheet.write_formula('G16', '=F16-H16', fmt_number_pink)
    worksheet.write_formula('K16', '=J16-L16', fmt_number_pink)
    worksheet.write_formula('O16', '=N16-P16', fmt_number_pink)
    #
    worksheet.merge_range('B18:E18', '=B16+D16+E16', fmt_number_center_border2) # Total Votes in Prior Elections
    worksheet.merge_range('F18:I18', '=F16+H16+I16', fmt_number_center_border2)
    worksheet.merge_range('J18:M18', '=J16+L16+M16', fmt_number_center_border2)
    worksheet.merge_range('N18:Q18', '=N16+P16+Q16', fmt_number_center_border2)
    #
    worksheet.write_formula('B24', '=SUM(B23:D23)', fmt_number_border)      # Total Active Voters in Prior Elections
    worksheet.write('C24', ' ', fmt_center_border)                     # border the blank cells
    worksheet.write('D24', ' ', fmt_center_border)
    worksheet.write_formula('F24', '=SUM(F23:H23)', fmt_number_border)
    worksheet.write('G24', ' ', fmt_center_border)                     # border the blank cells
    worksheet.write('H24', ' ', fmt_center_border)
    worksheet.write_formula('J24', '=SUM(J23:L23)', fmt_number_border)
    worksheet.write('K24', ' ', fmt_center_border)                     # border the blank cells
    worksheet.write('L24', ' ', fmt_center_border)
    worksheet.write_formula('N24', '=SUM(N23:P23)', fmt_number_border)
    worksheet.write('O24', ' ', fmt_center_border)                     # border the blank cells
    worksheet.write('P24', ' ', fmt_center_border)
    #
    worksheet.merge_range('B25:D25','=B18/B24', fmt_pct_center_pink)        # % Turnout in Prior Elections
    worksheet.merge_range('F25:H25', '=F18/F24', fmt_pct_center_pink)
    worksheet.merge_range('J25:L25', '=J18/J24', fmt_pct_center_pink)
    worksheet.merge_range('N25:P25', '=N18/N24', fmt_pct_center_pink)
    #
    worksheet.write_formula('B28', '=SUM(B25:N25)/4', fmt_pct_center_yel)   # Calculated Average Turnout
    worksheet.write_formula('C28', '=(F25+N25)/2', fmt_pct_center_yel)
    #
    worksheet.merge_range('F32:H32', '=B28', fmt_pct_center_yel)            # Avg T/O Estimate
    worksheet.merge_range('K32:M32', '=C28', fmt_pct_center_yel)
    #
    worksheet.merge_range('F33:H33', '=I8', fmt_number_center_org)          # Total Active Voters
    worksheet.merge_range('K33:M33', '=I8', fmt_number_center_org)
    #
    worksheet.merge_range('F34:H34', '=F32*F33', fmt_number_center_border)  # Expected Turnout
    worksheet.merge_range('K34:M34', '=K32*K33', fmt_number_center_border)
    #
    worksheet.merge_range('F36:H36', '=F34/2+100', fmt_number_center_grn)   # Estimated Votes Needed To Win
    worksheet.merge_range('K36:M36', '=K34/2+100', fmt_number_center_grn)
    #
    worksheet.write_formula('F40', '=B8', fmt_number_border)            # Avg Active Voters by Party
    worksheet.write_formula('G40', '=J8', fmt_number_border)
    worksheet.write_formula('H40', '=H8', fmt_number_border)
    worksheet.write_formula('K40', '=B8', fmt_number_border)
    worksheet.write_formula('L40', '=J8', fmt_number_border)
    worksheet.write_formula('M40', '=H8', fmt_number_border)
    #
    worksheet.write_formula('F41', '=F40*F32', fmt_number_border)       # Avg Turnout by Party
    worksheet.write_formula('G41', '=G40*F32', fmt_number_border)
    worksheet.write_formula('H41', '=H40*F32', fmt_number_border)
    worksheet.write_formula('K41', '=K40*K32', fmt_number_border)
    worksheet.write_formula('L41', '=L40*K32', fmt_number_border)
    worksheet.write_formula('M41', '=M40*K32', fmt_number_border)
    #
    worksheet.write_formula('F42', '=C4*C10+D4*D10+E4*E10+F4*F10+G4*G10', fmt_number_border)        # CONS Split of OTH
    worksheet.write_formula('H42', '=C5*C10+D5*D10+E5*E10+F10*F5+G5*G10', fmt_number_border)        # PROG Split of OTH
    worksheet.write_formula('K42', '=C4*C10+D4*D10+E4*E10+F4*F10+G4*G10', fmt_number_border)
    worksheet.write_formula('M42', '=C5*C10+D5*D10+E5*E10+F10*F5+G5*G10', fmt_number_border)
    #
    worksheet.write_formula('F43', '=F41+F42', fmt_number_border)       # Likely Outcome by Party
    worksheet.write_formula('G43', '=F43-H43', fmt_number_pink)
    worksheet.write_formula('H43', '=H41+H42', fmt_number_border)
    worksheet.write_formula('K43', '=K41+K42', fmt_number_border)
    worksheet.write_formula('L43', '=K43-M43', fmt_number_pink)
    worksheet.write_formula('M43', '=M41+M42', fmt_number_border)
    #
    #    E N D   S P R E A D S H E E T   F O R M U L A S
    #------------------------------------------------------
    #
    #  Now do values we load from the history and base files
    #
    #----------------------------------------------------------------------------------------
    #
    #     Current Cycle Registration and voting info from base.csv
    #
    REPreg = 0                                              # Registered Republicans
    DEMreg = 0                                              # Registered Democrats
    NPreg = 0                                               # Registered NLP
    IAPreg = 0                                              # Registered IAP
    LPreg = 0                                               # Registered LP
    GPreg = 0                                               # Registered GP
    OTHreg = 0                                              # registered OTH
    REPact = 0                                              # Active Republicans
    DEMact = 0                                              # Active Democrats
    NPact = 0                                               # Active NLP
    IAPact = 0                                              # Active IAP
    LPact = 0                                               # Actived LP
    GPact = 0                                               # Active GP
    OTHact = 0                                              # Active OTH
    MWRep = 0                                               # Moderate/Weak Republicans
    SOTH = 0                                                # Strong OTH
    cycledate = datetime(int(cycle),11,1)                   # max reg date for this cycle
    District = Dbase.to_dict(orient='list')                 # convert Dbase dataframe to dictionary with parallel columns
    #Dbase =[]                                               # release dataframe memory
    for x in range(0, baserows):
        status = District["Status"][x]                      # get ACTIVE/INACTIVE
        party = District["Party"][x]                        # get party registered to
        strength = District["LikelytoVote"][x]              # get STRONG/MODERATE/WEAK
#
#        regdate = datetime.strptime(row[bDict["RegDate"]],'%m/%d/%Y')     # get registration date as datetime object (really slow!)
#        if(regdate > cycledate):
#            continue                                        # ignore regs that couldn't have voted this cycle
        if (party == "Republican"):
            REPreg = REPreg + 1                             # count registered in this party
            if (status == "Active"):
                REPact = REPact + 1                         # count as active voter this party
            if ((strength == "MODERATE") or (strength == "WEAK")):
                MWRep = MWRep + 1                           # tally moderate or weak Republicans
        if (party == "Democrat"):
            DEMreg = DEMreg + 1                             # count registered in this party
            if (status == "Active"):
                DEMact = DEMact + 1                         # count as active voter this party
        if (party == "Non-Partisan"):
            NPreg = NPreg + 1                               # count registered in this party
            if (status == "Active"):
                NPact = NPact + 1                           # count as active voter this party
            if (strength == "STRONG"):
                SOTH = SOTH + 1                             # Count Strong voting Other
        if (party == "Independent American Party"):
            IAPreg = IAPreg + 1                             # count registered in this party
            if (status == "Active"):
                IAPact = IAPact + 1                         # count as active voter this party
            if (strength == "STRONG"):
                SOTH = SOTH + 1                             # Count Strong voting Other
        if (party == "Libertarian Party"):
            LPreg = LPreg + 1                               # count registered in this party
            if (status == "Active"):
                LPact = LPact + 1                           # count as active voter this party
            if (strength == "STRONG"):
                SOTH = SOTH + 1                             # Count Strong voting Other
        if (party == "Green Party"):
            GPreg = GPreg + 1                               # count registered in this party
            if (status == "Active"):
                GPact = GPact + 1                           # count as active voter this party
            if (strength == "STRONG"):
                SOTH = SOTH + 1                             # Count Strong voting Other
        if (party == "Other (All Others)"):
            OTHreg = OTHreg + 1                             # count registered in this party
            if (status == "Active"):
                OTHact = OTHact + 1                         # count as active voter this party
            if (strength == "STRONG"):
                SOTH = SOTH + 1                             # Count Strong voting Other
    #
    #  Now put calculated current registration values in their columns
    #
    worksheet.write('B7', REPreg, fmt_number)
    worksheet.write('B8', REPact, fmt_number)
    worksheet.write('C7', NPreg, fmt_number)
    worksheet.write('C8', NPact, fmt_number)
    worksheet.write('D7', IAPreg, fmt_number)
    worksheet.write('D8', IAPact, fmt_number)
    worksheet.write('E7', LPreg, fmt_number)
    worksheet.write('E8', LPact, fmt_number)
    worksheet.write('F7', GPreg, fmt_number)
    worksheet.write('F8', GPact, fmt_number)
    worksheet.write('G7', OTHreg, fmt_number)
    worksheet.write('G8', OTHact, fmt_number)
    worksheet.write('H7', DEMreg, fmt_number)
    worksheet.write('H8', DEMact, fmt_number)
    worksheet.write('B44', MWRep, fmt_number)
    worksheet.write('B45', SOTH, fmt_number)
    #
    #-------------------------------------------------------
    #  Get previous election cycle data from history files
    #
    #  Store 1st back election names & votes
    #
    # Note Info extracted to PnCandidates and PnVotes lists from load above
    #
    OVotes=0
    OName = ""                                                  # Assume no OTH candidate to start
    for x in range(4,8):                                        # See if there is an OTH candidate in this race
        if(P1Candidates != ""):
            OVotes = P1Votes[x]                                 # yes log their # votes
            OName = P1Candidates[x]                             # and last name (only take 1, if more we ignore)
            break
    worksheet.write('B15', P1Candidates[3], fmt_center_bold_border2L) # Republican Candidate last name
    if (len(P1Candidates[3]) > 8):
        worksheet.set_column('B:B', len(P1Candidates[3])*1.32)  # Handle Candidate name longer than 8 chars
    worksheet.write('C15', ' ', fmt_center_border)              # blank cell with border
    worksheet.write('D15', P1Candidates[2], fmt_center_bold_border)   # Democrat Candidate last name
    if (len(P1Candidates[2]) > 8):
        worksheet.set_column('D:D', len(P1Candidates[2])*1.32)  # Handle Candidate name longer than 8 chars
    worksheet.write('E15', OName, fmt_center_bold_border2R)     # OTH Candidate last name
    if (len(OName) > 8):
        worksheet.set_column('E:E', len(OName)*1.32)            # Handle Candidate name longer than 8 chars
    worksheet.write('B16', P1Votes[3],fmt_number_border2L)      # Republican Votes
    worksheet.write('D16', P1Votes[2],fmt_number_border)        # Democrat Votes
    worksheet.write('E16', OVotes,fmt_number_border2R)          # OTH Votes
    #
    #  Store 2nd back election names & votes
    #
    OVotes=0
    OName = ""
    for x in range(4,8):                                        # See if there is an OTH candidate in this race
        if(P2Candidates != ""):
            OVotes = P2Votes[x]
            OName = P2Candidates[x]
            break
    worksheet.write('F15', P2Candidates[3], fmt_center_bold_border2L)  # Republican Candidate last name
    if (len(P2Candidates[3]) > 8):
        worksheet.set_column('F:F', len(P2Candidates[3])*1.32)  # Handle Candidate name longer than 8 chars
    worksheet.write('G15', ' ', fmt_center_border)              # blank cell with border
    worksheet.write('H15', P2Candidates[2], fmt_center_bold_border)    # Democrat Candidate last name
    if (len(P2Candidates[2]) > 8):
        worksheet.set_column('H:H', len(P2Candidates[2])*1.32)  # Handle Candidate name longer than 8 chars
    worksheet.write('I15', OName, fmt_center_bold_border2R)     # OTH Candidate last name
    if (len(OName) > 8):
        worksheet.set_column('I:I', len(OName)*1.32)            # Handle Candidate name longer than 8 chars
    worksheet.write('F16', P2Votes[3],fmt_number_border2L)      # Republican Votes
    worksheet.write('H16', P2Votes[2],fmt_number_border)        # Democrat Votes
    worksheet.write('I16', OVotes,fmt_number_border2R)          # OTH Votes
    #
    #  Store 3rd back election names & votes
    #
    OVotes=0
    OName = ""
    for x in range(4,8):                                        # See if there is an OTH candidate in this race
        if(P3Candidates != ""):
            OVotes = P3Votes[x]
            OName = P3Candidates[x]
            break
    worksheet.write('J15', P3Candidates[3], fmt_center_bold_border2L) # Republican Candidate last name
    if (len(P3Candidates[3]) > 8):
        worksheet.set_column('J:J', len(P3Candidates[3])*1.32)  # Handle Candidate name longer than 8 chars
    worksheet.write('K15', ' ', fmt_center_border)              # blank cell with border
    worksheet.write('L15', P3Candidates[2], fmt_center_bold_border)   # Democrat Candidate last name
    if (len(P3Candidates[2]) > 8):
        worksheet.set_column('L:L', len(P3Candidates[2])*1.32)  # Handle Candidate name longer than 8 chars
    worksheet.write('M15', OName, fmt_center_bold_border2R)     # OTH Candidate last name
    if (len(OName) > 8):
        worksheet.set_column('M:M', len(OName)*1.32)            # Handle Candidate name longer than 8 chars
    worksheet.write('J16', P3Votes[3],fmt_number_border2L)      # Republican Votes
    worksheet.write('L16', P3Votes[2],fmt_number_border)        # Democrat Votes
    worksheet.write('M16', OVotes,fmt_number_border2R)          # OTH Votes
    #
    #  Store 4th back election names & votes
    #
    OVotes=0
    OName = ""
    for x in range(4,8):                                        # See if there is an OTH candidate in this race
        if(P4Candidates != ""):
            OVotes = P4Votes[x]
            OName = P4Candidates[x]
            break
    worksheet.write('N15', P4Candidates[3], fmt_center_bold_border2L)  # Republican Candidate last name
    if (len(P4Candidates[3]) > 8):
        worksheet.set_column('N:N', len(P4Candidates[3])*1.32)  # Handle Candidate name longer than 8 chars
    worksheet.write('O15', ' ', fmt_center_border)              # blank cell with border
    worksheet.write('P15', P4Candidates[2], fmt_center_bold_border)    # Democrat Candidate last name
    if (len(P4Candidates[2]) > 8):
        worksheet.set_column('P:P', len(P4Candidates[2])*1.32)  # Handle Candidate name longer than 8 chars
    worksheet.write('Q15', OName, fmt_center_bold_border2R)     # OTH Candidate last name
    if (len(OName) > 8):
        worksheet.set_column('Q:Q', len(OName)*1.32)            # Handle Candidate name longer than 8 chars
    worksheet.write('N16', P4Votes[3],fmt_number_border2L)      # Republican Votes
    worksheet.write('P16', P4Votes[2],fmt_number_border)        # Democrat Votes
    worksheet.write('Q16', OVotes,fmt_number_border2R)          # OTH Votes
    #
    #   --- For now put in blank bordered cells for undervote until we learn what we are doing there
    #
    worksheet.write('B17', ' ', fmt_center_border2L)
    worksheet.write('C17', ' ', fmt_center_border)
    worksheet.write('D17', ' ', fmt_center_border)
    worksheet.write('E17', ' ', fmt_center_border2R)
    worksheet.write('F17', ' ', fmt_center_border2L)
    worksheet.write('G17', ' ', fmt_center_border)
    worksheet.write('G17', ' ', fmt_center_border)
    worksheet.write('I17', ' ', fmt_center_border2R)
    worksheet.write('J17', ' ', fmt_center_border2L)
    worksheet.write('K17', ' ', fmt_center_border)
    worksheet.write('L17', ' ', fmt_center_border)
    worksheet.write('M17', ' ', fmt_center_border2R)
    worksheet.write('N17', ' ', fmt_center_border2L)
    worksheet.write('O17', ' ', fmt_center_border)
    worksheet.write('P17', ' ', fmt_center_border)
    worksheet.write('Q17', ' ', fmt_center_border2R)
    #
    #---------------------------------------------------------------------------------
    #
    #              History for Prior Election Turnout
    #
    P1found = 0
    P2found = 0
    P3found = 0
    P4found = 0
    extract = dfTO.loc[dfTO['County Name'] == Dist]                 # Get history for this district into dataframe
    for x in range(0, len(extract)):
        row = list(extract.iloc[x])                                 # Get next row of district history
        Dreg = row[2]                                               # Get DEM Registration for this year
        Rreg = row[9]                                               # Get REP Registration for this year
        Oreg = row[3] + row[4] + row[5] +row[6] +row[7] + row[8]    # total OTH Registration for this year
        if (row[0] == Prev1):
            P1found = 1
            worksheet.write('B23', Rreg,fmt_number_border)          # Republican Votes
            worksheet.write('C23', Oreg,fmt_number_border)          # OTH Votes
            worksheet.write('D23', Dreg,fmt_number_border)          # Democrat Votes
        if (row[0] == Prev2):
            P2found = 1
            worksheet.write('F23', Rreg,fmt_number_border)          # Republican Votes
            worksheet.write('G23', Oreg,fmt_number_border)          # OTH Votes
            worksheet.write('H23', Dreg,fmt_number_border)          # Democrat Votes
        if (row[0] == Prev3):
            P3found = 1
            worksheet.write('J23', Rreg,fmt_number_border)          # Republican Votes
            worksheet.write('K23', Oreg,fmt_number_border)          # OTH Votes
            worksheet.write('L23', Dreg,fmt_number_border)          # Democrat Votes
        if (row[0] == Prev4):
            P4found = 1
            worksheet.write('N23', Rreg,fmt_number_border)          # Republican Votes
            worksheet.write('O23', Oreg,fmt_number_border)          # OTH Votes
            worksheet.write('P23', Dreg,fmt_number_border)          # Democrat Votes
    if (P1found == 0):
        worksheet.write('B23', 0,fmt_number_border)                 # Republican Votes
        worksheet.write('C23', 10,fmt_number_border)                # OTH Votes
        worksheet.write('D23', 0,fmt_number_border)                 # Democrat Votes
    if (P2found == 0):
        worksheet.write('F23', 0,fmt_number_border)                 # Republican Votes
        worksheet.write('G23', 10,fmt_number_border)                # OTH Votes
        worksheet.write('G23', 0,fmt_number_border)                 # Democrat Votes
    if (P3found == 0):
        worksheet.write('N23', 0,fmt_number_border)                 # Republican Votes
        worksheet.write('O23', 10,fmt_number_border)                # OTH Votes
        worksheet.write('P23', 0,fmt_number_border)                 # Democrat Votes
    if (P4found == 0):
        worksheet.write('N23', 0,fmt_number_border)                 # Republican Votes
        worksheet.write('O23', 10,fmt_number_border)                # OTH Votes
        worksheet.write('P23', 0,fmt_number_border)                 # Democrat Votes
    #
    #  Writeout and close race spreadsheet file
    #
    try:
        workbook.close()
    except Exception as e:
        print('>>>>> Error Writing Spreadsheet file!!\n   Message, {m}'.format(m = str(e)))
    EndTime = time.time()
    print (f"Done! - Total Elapsed time is {int((EndTime - StartTime)*10)/10} seconds.\n")
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
