#************************************************************************************
#                                 nvvoter1.py                                       *
#                                                                                   *
#  Input:  Secretary of State Vote History .csv File (one vote per row)             *
#          Configuration Spreadsheet nvconfig.xlsx                                  *
#                                                                                   *
#  Output is a CSV file by voter of their voting history in selected elections      *
#                                                                                   *
# *********************************************************************************** 

import pandas as pd
import numpy as np
import sys, os
from os.path import isfile, join
import datetime
import time
import math
from io import BytesIO
#
#  Configuration SpreadSheet File Name, Header & Data Row Arrays
#
CfgFile = "nvconfig.xlsx"                       # program configuration spreadsheet
CfgHeadings =[]                                 # Array of Text Headings for spreadsheet
CfgRow = []                                     # Data from the Row of spreadsheet currently being processed

voterHistoryFile = "VoterList.VtHst.43842.060420175555.csv"
voterHistoryFileh = 0                           # Secretary of State Vote History File
voterHistoryLine = []

inputFile = ""
inputFileh = 0                                  # Secretary of State Eligible Voter File

baseFile = "base.csv"                           
baseFileh = 0
baseLine = []
baseProfile = []

debugFile = "debug.txt"
debugFileh = 0
debug = 0
duplicates = []

# list of email addresses to add
voterEmailFile = ""                             # name of file containing name and email cross reference
voterEmailFileh = 0
veframe = []
voterEmailArray = []
voterEmailHeadings = []
emailAdded = 0

stream = 0
voterDataHeading = ""
voterDataFile    = "voterdata-s.csv"
voterDataFileh = 0
voterData = 0
voterDataLine = []

electionValue = []

printFile = "print.txt"
printFileh = 0

helpReq   = 0
fileCount = 0

csvHeadings = []
line1Read = ''
linesRead = 0
printData = ""
linesWritten = 0
statsAdded   = 0
csvRowHash = []
stateVoterID = 0
date = 0
adjustedDate = 0
before = 0
vote = 0
cycle = 0
totalVotes      = 0
linesIncRead    = 0
linesIncWritten = 0
ignored         = 0
currentVoter = 0

voterDataHeading = ["statevoterid",
    "11/03/20 general",                     # index to here is 1 for configuration load
    "06/09/20 primary",
    "11/06/18 general",
    "06/12/18 primary",
    "11/08/16 general",
    "06/14/16 primary",
    "11/04/14 general",
    "06/10/14 primary",
    "11/06/12 general",
    "06/12/12 primary",
    "09/13/11 special",
    "11/02/10 general",
    "06/08/10 primary",
    "11/04/08 general",
    "08/12/08 primary",
    "11/07/06 general",
    "08/15/06 primary",
    "11/02/04 general",
    "09/07/04 primary",
    "06/03/03 special",
    "TotalVotes ",     #21 Calculated
    "Generals",        #22 Calculated
    "Primaries",       #23 Calculated
    "Polls",           #24 Calculated
    "Absentee",        #25 Calculated
    "Early",           #26 Calculated
    "Provisional",     #27 Calculated
    "LikelytoVote",    #28 Calculated
    "Score"]           #29 Calculated


precinctPolitical = 0
RegisteredDays   = 0
generalCount     = 0
primaryCount     = 0
pollCount        = 0
absenteeCount    = 0
provisionalCount = 0
earlyCount       = 0
activeVOTERS     = 0
activeREP        = 0
activeDEM        = 0
activeOTHR       = 0
totalVOTERS      = 0
totalGENERALS    = 0
totalPRIMARIES   = 0
totalPOLLS       = 0
totalABSENTEE    = 0
totalPROVISIONAL = 0
totalMAIL        = 0
totalSTR         = 0
totalMOD         = 0
totalWEAK        = 0
votesTotal       = 0
voterRank        = 0
voterScore       = 0
voterScore2      = 0
noVotes  = 0

voterHeadingDates = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
voterEarlyDates = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

#
#  Array to compile the highest voter ID that voted in each of the 20 elections being tracked
#
HighVoterID = [0,                                          # VoterID = 0 indicates this record
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]               # Highest Voter ID that voted in the 20 elections (inits at 0)
#
# Precinct file data
#
#  NumPct = Number of precincts (and thus the number of entries in each parallel precinct array)
#  PctPrecinct = Array of Precinct Numbers in this base.csv compilation
#  Pctxxx = parallel Array of counts for each item by precinct
#
pctFile = "precinct.csv"
pctFileh = 0
NumPct = 0                          # Number of Precincts
PctPrecinct =[]                     # Array of precinct numbers
PctCD = []                          # Array of Congressional Districts
PctAD =[]                           # Array of State Assembly Districts
PctSD =[]                           # Array of State Senate Districts
PctBoardofEd =[]                    # Array of Board of Ed
PctCntyComm =[]                     # Array of Board of Ed
PctRwards =[]                       # Array of Board of Ed
PctSwards =[]                       # Array of Board of Ed
PctSchBdTrust =[]                   # Array of Board of Ed
PctSchBdAtLrg =[]                   # Array of Board of Ed
PctGenerals =[]                     # Total general election votes this precinct
PctPrimaries =[]                    # Total primary election votes this precinct
PctPolls =[]                        # Total poll votes this precinct
PctAbsentee =[]                     # Total mail in votes this precinct
PctRegRep =[]                       # Array of # Registered Republicans
PctRegDem =[]                       # Array of # Registered Democrats
PctRegNP =[]                        # Array of # Registered Non-Partisans
PctRegIAP =[]                       # Array of # Registered Independent American Party
PctRegLP =[]                        # Array of # Registered Libertarian Party
PctRegGP =[]                        # Array of # Registered Green Party
PctRegOther =[]                     # Array of # Registered to Other Parties
PctStrongRep =[]                    # Array of # Strong Voting Republicans
PctModRep =[]                       # Array of # Moderate Voting Republicans
PctWeakRep =[]                      # Array of # Weak Voting Republicans
PctStrongDem =[]                    # Array of # Strong Voting Democrats
PctModDem =[]                       # Array of # Moderate Voting Democrats
PctWeakDem =[]                      # Array of # Weak Voting Democrats
PctStrongAllOther =[]               # Array of # Strong Voting All Other Parties
PctModAllOther =[]                  # Array of # Moderate Voting All Other Parties
PctWeakAllOther =[]                 # Array of # Weak Voting All Other Parties
PctActiveRep =[]                    # Array of # of active Republican
PctActiveDem =[]                     # Array of # of active Democrat
PctActiveAllOther =[]               # Array of # of active voter in All Other Parties
#
pctHeading = ""                     # Precint File Header
pctHeading =[   "Precinct",         # Precinct Number
                "CongDist",         # Congressional District
                "AssmDist",         # Assembly District
                "SenDist ",         # Senate District
                "BrdofEd",          # Board of education District
                "CntyComm",         # county commission 
                "Rwards",           # Reno wards
                "Swards",           # Sparks wards
                "SchBdTrust",       # Board of education trustes
                "SchBdAtLrg",       # Board of education at large
                "Generals",         # # General Election Votes over all Election cycles
                "Primaries",        # # Primary Election Votes Over All Election Cycles
                "Polls",            # # Voters Voting at pools Over all Election Cycles
                "Absentee",         # # Voters Voting by mail Over All Election Cycles
                "Reg-NP",           # Total Registered Non-Partisan
                "Reg-IAP",          # Total Registered Independent American Party
                "Reg-LP",           # Total Registered Libertarian Party
                "Reg-GP",           # Total Registered Green Party
                "Reg-Other",        # Total Registered Other (All Others)
                "Reg-Rep",          # Total Registered Republican
                "Active Rep",       # Republicans marked ACTIVE
                "% Rep",            # Percentage of registered Voters that are Republican
                "Reg-Dem",          # Total Registered Democrat
                "Active Dem",       # Democrats marked ACTIVE
                "% Dem",            # Percentage of registered Voters that are Democrat
                "Reg AllOther",     # Total Registered Other (All Others including NP, IAP, LP & GP)
                "Active AllOther",  # All Other Party voters marked ACTIVE
                "% AllOther",       # Percentage of reg Voters that are All Others including NP, IAP, LP & GP
                "#Strong Rep",      # Total Strong Voting Republicans
                "#Moderate Rep",    # Total Moderate Voting Republicans
                "#Weak Rep",        # Total Weak Voting Republicans
                "#Strong Dem",      # Total Strong Voting Democrats
                "#Moderate Dem",    # Total Moderate Voting Democrats
                "#Weak Dem",        # Total Weak Voting Democrats
                "#Strong Other",    # Total Strong Voting All Other Parties
                "#Moderate Other",  # Total Moderate All Other Parties
                "#Weak Other"]      # Total Weak All Other Parties

#                      base.csv header
fixedflds = 32;                     # 32 fixed fields before votedata
baseHeading = ["CountyID",     "StateID",  "Status",   "County",    "Precinct", "CongDist",
    "AssmDist",     "SenDist",  "BrdofEd",  "CntyComm",  "Rwards",   "Swards",   "SchBdTrust", "SchBdAtLrg",
    "First",        "Last",     "Middle",   "Suffix",    "Phone",    "email",
    "BirthDate",    "RegDate",  "Party",  "StreetNo",  "StreetName",   "Address1", "Address2", "City",
    "State",        "Zip",  "RegisteredDays", "Age", 
    "11/03/20-G",                         # index to here is 32
    "06/09/20-P",                         # these 20 election headers loaded from Config file
    "11/06/18-G",
    "06/12/18-P",   
    "11/08/16-G",
    "06/14/16-P",
    "11/04/14-G",
    "06/10/14-P",
    "11/06/12-G",
    "06/12/12-P",
    "09/13/11-S",
    "11/02/10-G",
    "06/08/10-P",
    "11/04/08-G",
    "08/12/08-P",
    "11/07/06-G",
    "08/15/06-P",
    "11/02/04-G",
    "09/07/04-P",
    "06/03/03-S",
    "TotalVotes", "Generals", "Primaries",
    "Polls",  "Absentee", 
    "Early",  "Provisional",
    "LikelytoVote", "Score"]
bDict = []

emailProfile = 0
emailHeading = ""
emailHeading = ["VoterID", "Precinct", "First", "Last", "Middle", "email"]

# email merge error report
emailLogFile = "email-adds-log.csv"
emailLogFileh = 0
emailLine = []

votingLine = ""
votingProfile = ""

precinct = "000000"
noVotes  = 0
noData   = 0


adPoliticalFile = "pctxref.csv"
politicalLine   = []
adPoliticalHash = []
adPoliticalHeadings = []
precinctPolitical
noPoliticalWarn = 0
Noxref = 0

ProgName = "NVVOTER"                  # Name of running program
#
#=====================================================================================================================
#
#************************************
#                                   *
# Print Log line to screen and file *
#                                   *
#************************************
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
    print ("py nvvoter.py -config <filename> -infile <filename> -regfile <filename> -outfile<filename>")
    print ("               -pctfile <filename> -xref <filename> -emailfile <filename>")
    print ("    -config <filename> NVVOTER Configuration .xlsx Spreadsheet")
    print ("    -infile <filename> Secretary of State VoterList.VtHst.nnn .csv File.")
    print ("    -regfile = Secretary Of State ElgbVtr file - default is infile with VtHst chgd to ElgbVtr")
    print ("    -outfile = compiled \"base\" file - default is base.csv")
    print ("    -pctfile = precinct summary file - default is precinct.csv")
    print ("    -xref = Precinct to political district cross reference file - default is pctxref.csv")
    print ("    -emailfile = optional file of email addresses to add to base.csv on name match\n")
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
    global CfgFile, voterHistoryFile, voterDataFile, baseFile
    global inputFile, voterEmailFile, pctFile, adPoliticalFile

    print("")
    hst = 0
    elg = 0
    NumParms = len(argv)                                            # get number of items in argument list
    if (NumParms != 0):
        x = 0                                                           # start with argument # 1
        while (x < NumParms):                                           # scan while any args left
            opt = argv[x]                                               # fetch parameter
            x += 1                                                      # bump index
            if (x < NumParms):
                arg = argv[x]                                           # parameter has arg, fetch it
            else:
                arg = ""                                                # parameter has no arg, probably -help
            x += 1                                                      # bump index
            if opt in ['-h', '-help', '-?']:
                printhelp()
                exit(2)
            elif opt == "-outfile":
                baseFile = arg
            elif opt == "-infile":
                voterHistoryFile = arg
                hst = 1
            elif opt == "-config":
                CfgFile = arg
            elif opt == "-regfile":
                inputFile = arg
                elg = 1
            elif opt == '-pctfile':
                pctFile = arg
            elif opt == '-xref':
                adPoliticalFile = arg
            elif opt == '-emailfile':
                voterEmailFile = arg
            else:
                print (f">>>UnKnown Option: {opt}\n")
                printhelp()                                     # Unknown Option - provide help and exit
                return(1)
    #
    #--------------------------------------------------------------
    #
    #  Try to handle the obvious single file defaults.  These are
    #    1. Wrong file type on either infile or -regfile option
    #    2. Missing -infile or -regfile option (only one specified)
    #
    #--------------------------------------------------------------
    #
    if ((hst == 1) and (elg == 0)):
        if ("elgbvtr" in voterHistoryFile.lower()):
            inputFile = voterHistoryFile                                        # ElgVtr is not history file, move to reg voter file
            voterHistoryFile = inputFile.lower().replace("elgbvtr","VtHst",1)   # alter history file name to VtHst
            elg = 1                                                             # say defaulted Eligible Voter File Name
        elif ("vthst" in voterHistoryFile.lower()):
            inputFile = voterHistoryFile.lower().replace("vthst","ElgbVtr",1)
            elg = 1                                                             # say defaulted Eligible Voter File Name
    if ((hst == 0) and (elg == 1)):
        if ("elgbvtr" in inputFile.lower()):                                    # only registered voter file specified
            voterHistoryFile = inputFile.lower().replace("elgbvtr","VtHst",1)   # alter history file name to VtHst
            hst = 1                                                             # say defaulted Vote History File Name
        elif ("vthst" in inputFile.lower()):
            voterHistoryFile = inputFile                                        # Specified History file for Registered Voter File, swap it
            inputFile = voterHistoryFile.lower().replace("vthst","ElgbVtr",1)   # alter name to ElgbVtr in inputFile
            hst = 1                                                             # say defaulted Vote History File Name
    #
    #------------------------------------------------------------------
    #
    #  Now check that we have both a Vote History and Eligible Voter File Name Given
    #
    if ((hst + elg) < 2):
        printLine (f">>>Must Specify -infile and/or -regfile, neither was given - aborting")
        return(1)
    printLine (f"Using {voterHistoryFile} as Vote History file")
    printLine (f"Using {inputFile} as Elibigle Voter file")
    #
    #  Fully Qualify input and config file names
    #
    Dir = os.getcwd()                                       # get our current working directory
    if(CfgFile[2] != ":"):
        CfgFile = os.path.join(Dir,CfgFile)                 # form fully qualified survey file name
        if (os.path.isfile(CfgFile) == False):
            printLine (f"Configuration File '{CfgFile}' does not exist...aborting\n")
            return(2)
    if(voterHistoryFile[2] != ":"):
        voterHistoryFile = os.path.join(Dir,voterHistoryFile)  # form fully qualified survey file name
        if (os.path.isfile(voterHistoryFile) == False):
            printLine (f"Voting History File '{voterHistoryFile}' does not exist...aborting\n")
            return(2)
    return (0)
#
#====================================================================================================
#
#---------------------------------------------------------------
#
#  Load the configuration spreadsheet.
#  Currently it only contains the election cycle dates, types and vote weights
#
def load_config():
    global electionValue, voterDataHeading, CfgFile, baseHeading
    #
    # load configuration file and load in the 20 election cycles to be used
    #
    printLine(f"Loading Configuration from {CfgFile} ...")
    if (CfgFile[-4:] == ".csv"):
        cframe = pd.read_csv (CfgFile,low_memory=False)         #  Read .csv config file into dataframe "cframe"
    else:
        cframe = pd.read_excel (CfgFile)                        #  Read .xls or .xlsx survey file into dataframe "cframe"
    CfgRows = len(cframe.index)                                 # get Number of Rows in Spreadsheet
    CfgCols = len(cframe.columns)                               # get Number of Columns in Spreadsheet
    printLine(f"... Configuration Loaded: {CfgRows} Data Rows of {CfgCols} columns each.")
    if (CfgCols < 3):
        printLine("Invalid Configuration, Headings Not:\n Election Date, Election Type, Vote Weight\n")
        return (2)
    #
    # Fetch and save Excel Header row text strings into CfgHeadings array
    #
    row = 0                                                     # row index for configuration spreadsheet
    for j in range(0,CfgRows):
        CfgHeadings = list(cframe.iloc[j])                      # get row data as a list
        row = j                                                 # save header row index once found
        if (isinstance(CfgHeadings[0],str)):
            if (CfgHeadings[0][0:1] == "#"):
                continue                                        # ignore comment lines before header row
        if (CfgHeadings[0] == "Election Date"):                 # Found Heading line, 
            if ((CfgHeadings[1] != "Election Type") or (CfgHeadings[2] != "Vote Weight")):
                printLine("Invalid Configuration, Headings Not:\n Election Date, Election Type, Vote Weight\n")
                return (2)
            break
    if (row >= CfgRows):
        printLine("Invalid Configuration, \"Election Date\" heading not found \n")
        return(2)
    #
    #  Now load the election date configuration data
    #
    ElecDates = []                                                  # electiondate array
    electionValue = []                                              # election value array
    edx = 0;                                                        # electiondate index
    for j in range(row+1,CfgRows):
        CfgRow = list(cframe.iloc[j])                               # read next config row
        if (isinstance(CfgRow[0],str)):
            if (CfgRow[0][0:1] == "#"):
                continue                                            # ignore comment lines
        rdate = CfgRow[0]                                           # Get Election Date
        if (isinstance(rdate,str)):
#            rdate = datetime.datetime.strptime(rdate,"%m/%d/%Y")             # date is string, convert to datetime object
            rdate = datetime.datetime(int(rdate[6:]), rdate(rdate[0:2]), int(rdate[3:5]),0 ,0 ,0)
        elif (isinstance(rdate,type(datetime.datetime.now())) != True):
            rdate = rdate.to_pydatetime()                           # date is pandas Timestamp, convert to datetime object
        yy="{0:4d}".format(rdate.year)                              # get yyyy
        yy = yy[2:]                                                 # get last two digits of year
        mm="{0:02d}".format(rdate.month)                            # get mm
        dd="{0:02d}".format(rdate.day)                              # get dd
        ElecDate = f"{mm}/{dd}/{yy} {CfgRow[1]}"                    # build "mm/dd/yy type" string
        ElecDates.append(ElecDate)                                  # save election column headers
        electionValue.append(CfgRow[2])                             # save election voting weights
        if (edx >= 19 ):
            break                                                   # only take in 20 elections
        edx += 1                                                    # count another election date scanned
    if (edx != 19):
        printLine(f"Invalid Election Date Configuration, must be 20 elections defined")
        return (2)
    printLine ("Configured to use these 20 elections")
    for j in range(0,20):
        printLine (f"{ElecDates[j]} Voting Weight={electionValue[j]}")   # display on console and in print log file
        voterDataHeading[j+1] = ElecDates[j]                        # copy to intermediate voter's vote history header
        baseHeading[j+fixedflds] = ElecDates[j]                     # and to base.csv header
    return (0)
#
#====================================================================================================
#
#---------------------------------------------------------------
#  routine: evaluateVoter
#
# determine if reliable voter by voting pattern over last five cycles
# tossed out special elections and mock elections
#  voter reg_date is considered
#  weights: strong, moderate, weak
# if registered < 2 years       gen >= 1 and pri <= 0   = STRONG
# if registered > 2 < 4 years   gen >= 1 and pri >= 0   = STRONG
# if registered > 4 < 8 years   gen >= 4 and pri >= 0   = STRONG
# if registered > 8 years       gen >= 6 and pri >= 0   = STRONG
#
def evaluateVoter():
    global totalGENERALS, totalPRIMARIES, totalPOLLS, totalABSENTEE
    global totalPROVISIONAL, totalSTR, totalMOD, totalWEAK
    global voterDataLine, voterDataHeading
    global generalCount, primaryCount, pollCount, absenteeCount, earlyCount
    global provisionalCount, votesTotal, voterRank, voterScore2

    generalPollCount  = 0           # init local variables
    generalEarlyCount = 0
    generalNotVote    = 0
    notElegible       = 0
    primaryPollCount  = 0
    primaryEarlyCount = 0
    primaryNotVote    = 0
    badcode           = 0
    badstring         = ""
    oldestCast        = 0
    #
    generalCount      = 0           # init global variables
    primaryCount      = 0
    pollCount         = 0
    absenteeCount     = 0
    earlyCount        = 0
    provisionalCount  = 0
    votesTotal        = 0
    voterRank         = ''
    for vote in range(1,21):
        badcode   = 1
        votecode = voterDataLine[vote]         # save bad entry in case no match

    # each election type is specified with its date - we only process primary/general

        if ("mock" in voterDataHeading[vote].lower()):
            badcode = 0
            continue                            # skip mock election
        if ("special" in voterDataHeading[vote].lower()):
            badcode = 0
            continue                            # skip special election
        if ("sparks" in voterDataHeading[vote].lower()):
            badcode = 0
            continue                            # skip sparks election
        #
        # record a general vote
        # if there is no vote recorded (shown with a "blank") then NOT ELEGIBLE
        #
        if ("general" in voterDataHeading[vote].lower()):
            if ((votecode == ' ') or  (votecode == "" )):
                badcode = 0
                notElegible += 1
                continue
            #
            # the following vote codes are supported
            # - EV early vote
            # - FW federal write in
            # - MB mail ballot
            # - PP polling place
            # - PV provisional vote
            # - BR ballot received (prior to election day, becomes MB at election time)
            #
            if (votecode == 'N' ):
                badcode = 0
                generalNotVote += 1
                continue
            if (votecode == 'PP'):
                generalPollCount += 1
                generalCount     += 1
                pollCount        += 1
                votesTotal       += 1
                oldestCast = vote
                badcode    = 0
                continue
            if (votecode == 'FW'):
                generalPollCount += 1
                generalCount     += 1
                pollCount        += 1
                votesTotal       += 1
                oldestCast = vote
                badcode    = 0
                continue
            if (votecode == 'EV' ):
                generalEarlyCount += 1
                earlyCount        += 1
                generalCount      += 1
                votesTotal        += 1
                oldestCast = vote
                badcode    = 0
                continue
            if (votecode == 'MB' ):
                generalEarlyCount += 1
                generalCount      += 1
                earlyCount        += 1
                absenteeCount     += 1
                votesTotal        += 1
                oldestCast = vote
                badcode    = 0
                continue
            if (votecode == 'PV' ):
                generalCount      += 1
                provisionalCount  += 1
                votesTotal        += 1
                oldestCast = vote
                badcode    = 0
                continue
            if (votecode == 'BR' ):
              #  generalCount     += 1
              #  provisionalCount += 1
              #  votesTotal       += 1
                oldestCast = vote
                badcode    = 0
                continue
            if (badcode != 0 ):
                printLine(f"Unknown General Election Code {badstring} for voter {currentVoter}")
                badcode = 0
        #
        # record a primary vote
        # if there is no vote recorded shown with a "blank" then NOT ELEGIBLE
        #
        if ("primary" in voterDataHeading[vote].lower()):
            if (votecode ==  ' ' ):
                notElegible += 1
                badcode = 0
                continue
            if (votecode == "" ):
                notElegible += 1
                badcode = 0
                continue
            if (votecode == 'N' ):
                primaryNotVote += 1
                badcode = 0
                continue
            if (votecode == 'PP' ):
                primaryPollCount += 1
                primaryCount     += 1
                pollCount        += 1
                votesTotal       += 1
                oldestCast = vote
                badcode    = 0
                continue
            if (votecode == 'EV' ):
                primaryEarlyCount += 1
                earlyCount        += 1
                primaryCount      += 1
                votesTotal        += 1
                oldestCast = vote
                badcode    = 0
            if (votecode == 'MB' ):
                primaryEarlyCount += 1
                primaryCount      += 1
                earlyCount        += 1
                absenteeCount     += 1
                votesTotal        += 1
                oldestCast = vote
                badcode    = 0
                continue
            if (votecode == 'BR' ):
               # primaryEarlyCount += 1
               # primaryCount      += 1
               # earlyCount        += 1
               # absenteeCount     += 1
               # votesTotal        += 1
                oldestCast = vote
                badcode    = 0
                continue
            if (votecode == 'PV' ):
                primaryCount     += 1
                provisionalCount += 1
                votesTotal       += 1
                oldestCast = vote
                badcode    = 0
                continue
            if (badcode != 0 ):
                printLine(f"Unknown Primary Election Code {badstring} for voter {currentVoter}")
                badcode = 0
        if (badcode != 0 ):
            printLine(f"Unknown Vote Code {badstring} for voter {currentVoter}")
            badcode = 0
    #
    # Likely to vote score:
    # if registered < 2 years       gen <= 1 || notelig >= 1            = WEAK
    # if registered < 2 years       gen == 1 ||                         = MODERATE
    # if registered < 2 years       gen == 2 ||                         = STRONG

    # if registered > 2 < 4 years   gen <= 0 || notelig >= 1            = WEAK
    # if registered > 2 < 4 years   gen >= 2 && pri >= 0                = MODERATE
    # if registered > 2 < 4 years   gen >= 3 && pri >= 1                = STRONG

    # if registered > 4 < 8 years   gen >= 0 || notelig >= 1            = WEAK
    # if registered > 4 < 8 years   gen >= 0 && gen <= 2  and pri == 0  = WEAK
    # if registered > 4 < 8 years   gen >= 2 && gen <= 5  and pri >= 0  = MODERATE
    # if registered > 4 < 8 years   gen >= 3 && gen <= 12 and pri >= 0  = STRONG

    # if registered > 8 years   gen >= 0 && gen <= 2 || notelig >= 1    = WEAK
    # if registered > 8 years   gen >= 0 && gen <= 4  and pri == 0      = WEAK
    # if registered > 8 years   gen >= 3 && gen <= 9  and pri >= 0      = MODERATE
    # if registered > 8 years   gen >= 6 && gen <= 12 and pri >= 0      = STRONG
    #
    if (votesTotal > 0):
        voterScore  = ((generalCount + primaryCount ) / oldestCast) * 10
        voterScore2 = round(voterScore)
    else:
        voterScore2 = 0                               # if never voted, call WEAK
    #
    # voted, get weak, moderate or string
    if (voterScore2 > 5):
        voterRank = "STRONG"
        totalSTR += 1
    if ((voterScore2 >= 3) and (voterScore2 <= 5)):
        voterRank = "MODERATE"
        totalMOD += 1
    if (voterScore2 < 3):
        voterRank = "WEAK"
        totalWEAK += 1
    #
    totalGENERALS    = totalGENERALS + generalCount
    totalPRIMARIES   = totalPRIMARIES + primaryCount
    totalPOLLS       = totalPOLLS + pollCount
    totalABSENTEE    = totalABSENTEE + absenteeCount
    totalPROVISIONAL = totalPROVISIONAL + provisionalCount
    return
#
#====================================================================================================
#
#   Numeric Binary Search
#
# index = binarylookup(list, val)
#   list = is a sorted list of numeric values
#   val = is the target value that might be in the list.
#
#   binarylookup() returns the list index such that list[index] = val
#                  returns -1 if val not in list
#
def binarylookup (list, val):
#    d = 0
#    if (len(list) == NumPct):
#        d = 1
#        print (f"Searching {len(list)} entries  for {val}")
    high = len(list) - 1                        # set index to last entry in list
    low = 0                                     # set low to 1st entry of list
    t = 0                                       # define look index
    var = 0                                     # define local temp variable
    while (low <= high):                        # While the window is open
        t = int( (low + high ) / 2 )            # Try the middle element
        var = list[t]
        if (var < val):                         # Raise bottom
            low  = t + 1                        # to 1 above this entry
            continue
        if (var > val):                         # Lower top
            high = t - 1                        # to 1 below this entry
            continue
#        if (d != 0):
#            print (f"Found {val} at index {t}")
        return (t)                              # We've found val!
#    if(d !=0):
#        print (f" Didn't Find {val}")
    return (-1)                                 # The word isn't there.
#
#====================================================================================================
#
#  Force item to be type string
#
def makestr(temp):
    if(isinstance(temp,str)):
        return(temp)
    if(isinstance(temp,float)):
        temp = str(int(temp))
        return (temp)
    temp = (str(temp))
    return (temp)
#
#====================================================================================================
#
#---------------------------------------------------------------
#
#  Add a new precinct row to the parallel precinct tables
#  keepin the tables in ascending precinct order
#
def add_pct():
    global NumPct, baseLine, PctPrecinct, PctCD, PctAD, PctSD, PctBoardofEd
    global PctCntyComm, PctRwards, PctSwards, PctSchBdTrust, PctSchBdAtLrg
    global PctGenerals, PctPrimaries, PctPolls, PctAbsentee, PctRegRep
    global PctRegDem, PctRegNP, PctRegIAP, PctRegLP, PctRegGP, PctRegOther
    global PctStrongRep, PctModRep, PctWeakRep, PctStrongDem, PctModDem
    global PctWeakDem, PctStrongAllOther, PctModAllOther, PctWeakAllOther
    global PctActiveRep, PctActiveDem, PctActiveAllOther, bDict

#    newpct =  baseLine[bDict["Precinct"]]
    #
    #  Append a row, as we know we need a new one
    #
    PctPrecinct.append(int(baseLine[bDict["Precinct"]]))    # set precinct number
    PctCD.append(baseLine[bDict["CongDist"]])               # set CD for this precinct
    PctAD.append(baseLine[bDict['AssmDist']])               # set Assembly District
    PctSD.append(baseLine[bDict['SenDist']])                # set Senate District
    PctBoardofEd.append(baseLine[bDict['BrdofEd']])         # set Board of Education
    PctCntyComm.append(baseLine[bDict['CntyComm']])         # set Board of Education
    PctRwards.append(baseLine[bDict['Rwards']])             # set Board of Education
    PctSwards.append(baseLine[bDict['Swards']])             # set Board of Education
    PctSchBdTrust.append(baseLine[bDict['SchBdTrust']])     # set Board of Education
    PctSchBdAtLrg.append(baseLine[bDict['SchBdAtLrg']])     # set Board of Education
    PctGenerals.append(0)                                   # init rest of row's data to zeroes
    PctPrimaries.append(0)
    PctPolls.append(0)
    PctAbsentee.append(0)
    PctRegRep.append(0)                                     # init rest of row's data to zeroes
    PctRegDem.append(0)
    PctRegNP.append(0)
    PctRegIAP.append(0)
    PctRegLP.append(0)
    PctRegGP.append(0)
    PctRegOther.append(0)
    PctStrongRep.append(0)
    PctModRep.append(0)
    PctWeakRep.append(0)
    PctStrongDem.append(0)
    PctModDem.append(0)
    PctWeakDem.append(0)
    PctStrongAllOther.append(0)
    PctModAllOther.append(0)
    PctWeakAllOther.append(0)
    PctActiveRep.append(0)
    PctActiveDem.append(0)
    PctActiveAllOther.append(0)
    if (NumPct == 0):
        NumPct = NumPct+1;                                  # say we added an array row
#        print (f"Appended {newpct}")
#        print (PctPrecinct)
        return                                              # first entry, no more to do
    #
    #  Maintain the list in sorted order.
    #
    ix = 0
    newpct = int(baseLine[bDict["Precinct"]])               # precinct we're adding
    while(PctPrecinct[ix] <= newpct):                        # find where it goes in the list
        ix = ix+1
        if ix > NumPct:
            NumPct = NumPct+1;                              # say we added an array row
#            print (f"Appended {newpct}")
#            print (PctPrecinct)
            return                                          # we got lucky it went at the end
    #print (f"NumPct = {NumPct}, ix = {ix}, newpct={newpct}")
    #print (f"P[ix-1]={PctPrecinct[ix-1]}, p[ix]={PctPrecinct[ix]}, P[ix+1]={PctPrecinct[ix+1]}")
    #
    #  Doesn't go at the end, ix points to insertion row.  Next we have to move
    #  all items from row ix to row NumPct-1 up one slot to make room for insertion
    #
    for j in range(NumPct-1,ix-1,-1):                           # j counts from NumPct-1 down to (and including) ix
        PctPrecinct[j+1] = PctPrecinct[j]                       # move up precinct number
        PctCD[j+1] = PctCD[j]                                   # move up CD for this precinct
        PctAD[j+1] = PctAD[j]                                   # move up Assembly District
        PctSD[j+1] = PctSD[j]                                   # move up Senate District
        PctBoardofEd[j+1] =PctBoardofEd[j]                      # move up Board of Education
        PctCntyComm[j+1] = PctCntyComm[j]                       # move up Board of Education
        PctRwards[j+1] = PctRwards[j]                           # move up Board of Education
        PctSwards[j+1] = PctSwards[j]                           # move up Board of Education
        PctSchBdTrust[j+1] = PctSchBdTrust[j]                   # move up Board of Education
        PctSchBdAtLrg[j+1] = PctSchBdAtLrg[j]                   # move up Board of Education
        PctGenerals[j+1] = PctGenerals[j]                       # move up rest of this row's data
        PctPrimaries[j+1] = PctPrimaries[j]
        PctPolls[j+1] = PctPolls[j]
        PctAbsentee[j+1] = PctAbsentee[j]
        PctRegRep[j+1] = PctRegRep[j]
        PctRegDem[j+1] = PctRegDem[j]
        PctRegNP[j+1] = PctRegNP[j]
        PctRegIAP[j+1] = PctRegIAP[j]
        PctRegLP[j+1] = PctRegLP[j]
        PctRegGP[j+1] = PctRegGP[j]
        PctRegOther[j+1] = PctRegOther[j]
        PctStrongRep[j+1] = PctStrongRep[j]
        PctModRep[j+1] = PctModRep[j]
        PctWeakRep[j+1] = PctWeakRep[j]
        PctStrongDem[j+1] = PctStrongDem[j]
        PctModDem[j+1] = PctModDem[j]
        PctWeakDem[j+1] = PctWeakDem[j]
        PctStrongAllOther[j+1] = PctStrongAllOther[j]
        PctModAllOther[j+1] = PctModAllOther[j]
        PctWeakAllOther[j+1] = PctWeakAllOther[j]
        PctActiveRep[j+1] = PctActiveRep[j]
        PctActiveDem[j+1] = PctActiveDem[j]
        PctActiveAllOther[j+1] = PctActiveAllOther[j]
    #
    #  ix points to open space we opened up.  Now do insertions there
    #
    PctPrecinct[ix] = int(baseLine[bDict["Precinct"]])      # set precinct number
    PctCD[ix] = baseLine[bDict["CongDist"]]                 # set CD for this precinct
    PctAD[ix] = baseLine[bDict['AssmDist']]                 # set Assembly District
    PctSD[ix] = baseLine[bDict['SenDist']]                  # set Senate District
    PctBoardofEd[ix] = baseLine[bDict['BrdofEd']]           # set Board of Education
    PctCntyComm[ix] = baseLine[bDict['CntyComm']]           # set Board of Education
    PctRwards[ix] = baseLine[bDict['Rwards']]               # set Board of Education
    PctSwards[ix] = baseLine[bDict['Swards']]               # set Board of Education
    PctSchBdTrust[ix] = baseLine[bDict['SchBdTrust']]       # set Board of Education
    PctSchBdAtLrg[ix] = baseLine[bDict['SchBdAtLrg']]       # set Board of Education
    PctGenerals[ix] = 0                                     # init rest of row's data to zeroes
    PctPrimaries[ix] = 0
    PctPolls[ix] = 0
    PctAbsentee[ix] = 0
    PctRegRep[ix] = 0                                       # init rest of row's data to zeroes
    PctRegDem[ix] = 0
    PctRegNP[ix] = 0
    PctRegIAP[ix] = 0
    PctRegLP[ix] = 0
    PctRegGP[ix] = 0
    PctRegOther[ix] = 0
    PctStrongRep[ix] = 0
    PctModRep[ix] = 0
    PctWeakRep[ix] = 0
    PctStrongDem[ix] = 0
    PctModDem[ix] = 0
    PctWeakDem[ix] = 0
    PctStrongAllOther[ix] = 0
    PctModAllOther[ix] = 0
    PctWeakAllOther[ix] = 0
    PctActiveRep[ix] = 0
    PctActiveDem[ix] = 0
    PctActiveAllOther[ix] = 0
    NumPct = NumPct + 1                                   # say row was added
#    print (f"Inserted {newpct} at index={ix}")
#    print (PctPrecinct)
    return
#
#====================================================================================================
#
#---------------------------------------------------------------
#
#  Write out the precinct matrix data to the precinct.csv file
#
def write_precinct():
    global PctPrecinct, NumPct, PctRegNP, PctRegIAP, PctRegLP, PctRegGP, PctRegOther
    global PctRegRep, PctRegDem, pctFileh, PctBoardofEd, PctCntyComm, PctRwards
    global PctSwards, PctSchBdTrust, PctSchBdAtLrg, PctGenerals, PctPrimaries, PctPolls
    global PctAbsentee, PctActiveRep, PctActiveDem, PctActiveAllOther, PctStrongRep
    global PctModRep, PctWeakRep, PctStrongDem, PctModDem, PctWeakDem
    global PctStrongAllOther, PctModAllOther, PctWeakAllOther

    lineout = ""
    totvote = 0
    pctRep = 0
    pctDem = 0
    pctAllOther = 0
    numAllOther = 0
    j = 0
    i = 0
    PctSort = []
    PctSort = sorted(PctPrecinct)                                                   # copy a list of precinct numbers sorted in ascending order
    #
    #  Write out unsorted list of data in sorted precinct order
    #
    for j in range(0,NumPct):                                                       # this loops makes us write all rows out
        i = 0
        for z in range(0,NumPct):                                                   # this loop makes the output sorted
            if (PctPrecinct[i] == PctSort[j]):
                break 
            i += 1                                                                  # i points to next ascending precinct
        # calc  voters registereed to all other parties in precinct
        numAllOther = PctRegNP[i] + PctRegIAP[i] + PctRegLP[i] + PctRegGP[i] + PctRegOther[i]
        totvote = PctRegRep[i] + PctRegDem[i] + numAllOther                         # Calc Total Voters in precinct
        if (totvote == 0):
            totvote = 1                                                             # avoid divide by zero if no voters in a precinct
        pctRep = "{0:5.2f}".format( ((PctRegRep[i] / totvote * 10000)+.5)/100 )     # percent of precinct republican 
        pctDem = "{0:5.2f}".format( ((PctRegDem[i] / totvote * 10000)+.5)/100 )     # percent of precinct democrat
        pctAllOther = "{0:5.2f}".format( ((numAllOther / totvote * 10000)+.5)/100 ) # percent of precinct All Other Party Registration
        #
        #  There's probably a better way to build the output line, but I don't know
        #  what it is so here goes brute force.
        #
        lineout = makestr(PctPrecinct[i]) + "," + PctCD[i] + "," + PctAD[i] + "," + PctSD[i] + "," 
        lineout = lineout + PctBoardofEd[i] + "," + PctCntyComm[i] + "," + PctRwards[i] + ","
        lineout = lineout + PctSwards[i] + "," + PctSchBdTrust[i] + "," + PctSchBdAtLrg [i] + ","
        lineout = lineout + makestr(PctGenerals[i]) + "," + makestr(PctPrimaries[i]) + ","
        lineout = lineout + makestr(PctPolls[i]) + "," + makestr(PctAbsentee[i]) + "," + makestr(PctRegNP[i]) + ","
        lineout = lineout + makestr(PctRegIAP[i]) + "," + makestr(PctRegLP[i]) + "," + makestr(PctRegGP[i]) + ","  + makestr(PctRegOther[i]) + ","
        lineout = lineout + makestr(PctRegRep[i]) + "," + makestr(PctActiveRep[i]) + "," + pctRep + "%,"
        lineout = lineout + makestr(PctRegDem[i]) + "," + makestr(PctActiveDem[i]) + "," + pctDem + "%,"
        lineout = lineout + makestr(numAllOther) + "," + makestr(PctActiveAllOther[i]) + "," + pctAllOther + "%,"
        lineout = lineout + makestr(PctStrongRep[i]) + "," + makestr(PctModRep[i]) + "," + makestr(PctWeakRep[i]) + ","
        lineout = lineout + makestr(PctStrongDem[i]) + "," + makestr(PctModDem[i]) + "," + makestr(PctWeakDem[i]) + ","
        lineout = lineout + makestr(PctStrongAllOther[i]) + "," + makestr(PctModAllOther[i]) + "," + makestr(PctWeakAllOther[i])

        print (lineout, file=pctFileh)
    return
#
#====================================================================================================
#
#---------------------------------------------------------------
#
#  Calculate data for precinct.csv file
#  Called for each line in S.O.S. data file after
#  all processing to create $baseLine Hash for this record
#
def calc_precinct():
    global bDict, baseLine, PctPrecinct, NumPct, PctGenerals
    global PctPrimaries, PctPolls, PctAbsentee
    global PctRegRep, PctActiveRep, PctStrongRep, PctModRep, PctWeakRep
    global PctRegDem, PctActiveDem, PctStrongDem, PctModDem, PctWeakDem
    global PctRegIAP, PctRegDem, PctRegLP, PctRegNP, PctRegOther
    global PctActiveAllOther, PctStrongAllOther, PctModAllOther, PctWeakAllOther
    i =0
    Active=0
    if (NumPct == 0):
        add_pct()
        i=0                                        #1st precinct added, set index
    else:
        i = binarylookup(PctPrecinct,int(baseLine[bDict["Precinct"]]))
        if (i == -1):
            i = NumPct                              # this precinct not in list, this will be its index
            add_pct()                               # new precinct, add a row for it
    #
    #  i = index for this precinct's row in the precinct parallel array matrix
    #
    #  Accumulate the stats from this voter's $baseLine data.
    #
    PctGenerals[i] = PctGenerals[i] + int(baseLine[bDict["Generals"]])
    PctPrimaries[i] = PctPrimaries[i] + int(baseLine[bDict["Primaries"]])
    PctPolls[i] = PctPolls[i] + int(baseLine[bDict["Polls"]])
    PctAbsentee[i] = PctAbsentee[i] + int(baseLine[bDict["Absentee"]])
    Active=0                                        # Assume Inactive Voter
    if (baseLine[bDict["Status"]] == "Active"):
        Active = 1                                  # set 1 more Active Voter
    if (baseLine[bDict["Party"]] == "Republican"):
        #
        #  process Republican Voter
        #
        PctRegRep[i] += 1                           # Count another Registered Republican
        PctActiveRep[i] = PctActiveRep[i] + Active  # accumulate # active republican voters in precinct
        if (baseLine[bDict["LikelytoVote"]] == "STRONG"):
            PctStrongRep[i] += 1                    # Count as strong republican
            return
        if (baseLine[bDict["LikelytoVote"]] == "MODERATE"):
            PctModRep[i] += 1                       # Count as moderate republican
            return
        if (baseLine[bDict["LikelytoVote"]] == "WEAK"):
            PctWeakRep[i] +=1                       # Count as weak republican
            return
        if (baseLine[bDict["LikelytoVote"]] == "NEVER"):
            PctWeakRep[i] +=1                       # Count as weak republican
            return
        print (f"Error: Registered Republican w/unknown LikelytoVote")
        print (baseLine)
        exit (0)
        return                                      # done with this voter
    if (baseLine[bDict["Party"]] == "Democrat"):
        #
        #  process Democrat Voter
        #
        PctRegDem[i] += 1                           # Count another Registered Democrat
        PctActiveDem[i] = PctActiveDem[i] + Active  # accumulate # active Democrat voters in precinct
        if (baseLine[bDict["LikelytoVote"]] == "STRONG"):
            PctStrongDem[i] += 1                    # Count as strong Democrat
            return
        if (baseLine[bDict["LikelytoVote"]] == "MODERATE"):
            PctModDem[i] += 1                       # Count as moderate Democrat
            return
        if (baseLine[bDict["LikelytoVote"]] == "WEAK"):
            PctWeakDem[i] += 1                      # Count as weak Democrat
            return
        if (baseLine[bDict["LikelytoVote"]] == "NEVER"):
            PctWeakDem[i] +=1                       # Count as weak republican
            return
        print (f"Error: Registered Democrat w/unknown LikelytoVote")
        print (baseLine)
        exit (0)
        return                                      # done with this voter
    #
    #  Voter is not Republican or Democrat, so do the All OTHER PARTY stats\
    #
    PctActiveAllOther[i] = PctActiveAllOther[i] + Active # accumulate # active All non dem or Rep Party voters in precinct
    if (baseLine[bDict["LikelytoVote"]] == "STRONG"):
        PctStrongAllOther[i] += 1                   # Count as strong Other
    if (baseLine[bDict["LikelytoVote"]] == "MODERATE"):
        PctModAllOther[i] += 1                      # Count as moderate Other
    if (baseLine[bDict["LikelytoVote"]] == "WEAK"):
        PctWeakAllOther[i] += 1                     # Count as weak Other
    if (baseLine[bDict["LikelytoVote"]] == "NEVER"):
        PctWeakAllOther[i] +=1                      # Count as weak republican
    #
    #  Now Try to Find which OTHER party we might care about
    #
    party = baseLine[bDict["Party"]]
    if (party == "Independent American Party"):
        PctRegIAP[i] += 1
        return
    if (party == "Green Party"):
        PctRegGP[i] += 1
        return
    if (party == "Non-Partisan"):
        PctRegNP[i] += 1
        return
    if (party == "Libertarian Party"):
        PctRegLP[i] += 1
        return
    PctRegOther[i] += 1                            # Count as Registered some Other Party
    return
#
#====================================================================================================
#
#*************************************************************
#       >>>>>  M A I N   P R O G R A M   S T A R T  <<<<<    *
#*************************************************************
#
def main():
    #
    ##############################################################################
    #  Define all of the global tables and variables main program needs access to
    #
    global CfgFile, voterHistoryFile, voterDataFile, printFileh, ProgName
    global voterDataHeading, voterHeadingDates, voterEarlyDates, stateVoterID
    global ignored, totalVotes, votesTotal, generalCount, primaryCount, pollCount
    global absenteeCount, earlyCount, provisionalCount, voterRank, voterScore2
    global linesWritten, linesIncWritten, voterDataLine, stream, linesIncRead
    global inputFile, voterEmailFile, pctFile, adPoliticalFile, baseFile, baseFileh
    global Noxref, noPoliticalWarn,duplicates, noVotes, noData, statsAdded
    global baseLine, bDict, pctFileh, emailAdded
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
    ProgName = sys.argv[0][0:-3].upper()            # Stash program Name (minus .py) in upper case for PrintLine
    ec = args(sys.argv[1:])                         # Get command line arguments if any
    if (ec != 0):
        return (ec)                                 # get out if error
    #
    ec = load_config()                              # load configuration spreadsheet
    if (ec != 0):
        return (ec)                                 # get out if error
    #
    #  initialize oldest election date we care about
    #
    temp = voterDataHeading[20]                                 # Fetch Oldest Election we're configured for
    oldestElection = temp[0:8]                                  # extract date portion of string
    oldestDate = datetime.datetime.strptime(oldestElection,"%m/%d/%y")   # date is string, convert to datetime object
    printLine(f"Oldest Election Date to use: {oldestElection}")        # display to logging files(s)
    #
    # initialize binary election date/time object arrays from configuration test dates
    #
    for vote in range(1,21):
        edate  = voterDataHeading[vote][0:8]
        electiondate = datetime.datetime.strptime(edate,"%m/%d/%y")      # get this election date as datetime object
        voterHeadingDates[vote] = electiondate;                 # this is election date
        voterEarlyDates[vote]   = (electiondate - datetime.timedelta(days=14) )   # this is early voting start
    #
    #  At this point:
    #     1. voterDataHeading[1-20]  contain text election date & type
    #     2. voterHeadingDates[1-20] contain datetime object election dates
    #     3. voterEarlyDates[1-20]   contain datetime object early voting start
    #
    #
    #  Configuration loaded, Open Secretary of State Vote History .csv File
    #
    printLine(f"Loading Vote History File: {voterHistoryFile}. ")
    if (voterHistoryFile[-4:] == ".csv"):
        hframe = pd.read_csv (voterHistoryFile,low_memory=False)    #  Read .csv Vote History file into dataframe "hframe"
    else:
        hframe = pd.read_excel (voterHistoryFile)                   #  Read .xls or .xlsx Vote History file into dataframe "hframe"
    csvHeadings = list(hframe.columns)                              #  Get Vote History Headers
    hfrows = len(hframe.index)
    #
    #   Sort the dataframe on Voter ID so output will be sorted in that order
    #
    printLine(f"Sorting Vote History File on Voter ID...")
    hframe.sort_values(by=['VoterID'], inplace=True)
    hDict = hframe.to_dict(orient='list')                           # convert hframe to dictionary with parallel columns
    hframe =[]                                                      # release dataframe memory
    printLine("Building Voting History for Configured Elections from {0:,} votes...".format(hfrows))
    #
    #  Open memory stream and write output file header line
    #
    stream = BytesIO()                                  # open stream
    vhcols = voterDataHeading                           # get copy of elections header for acculmulator columns
    for i in range(0,len(vhcols)):
        vhcols[i] = f"\"{vhcols[i]}\""                  # Quote Header Fields to allow blanks in Titles
    temp = ",".join(vhcols) + "\n"                      # join into comma separated string with newline               
    stream.write((temp).encode('ASCII'))                # write header row to memory stream
    #
    #   Now process individual vote records into combined vote by voter records
    # 
    #
    #- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    # 
    #  Combine individual vote history records into a single record per voter
    #  listing vote codes in configured elections or null if no vote in that election.
    #
    #    - currentVoter = record-id
    #    - get record from voterHistoryFileh
    #    - if currentVoter is same as stateVoterID and Election is one configured, add vote history to row
    #    - after all vote history records read for a voter, create calculated values and write row to output stream
    #      stateVoterID = currentVoter
    #   endloop
    #
    #   after all Voter History Records processed, convert CSV data stream to DataFrame for next process
    #
    #- - - - - - - - - - - - 
    #
    CycleCols = range(1,21)                             # loop iterator for # of election cycles
    VHdrCols = range(0,len(voterDataHeading))           # loop iterator for # output columns
    voterDataLine =  [""] * len(voterDataHeading)       # build list of null strings for accumulator of by voter "row"
    #
    stateVoterID  = hDict["VoterID"][0]                 # set first record Voter ID as first to accumulate
    row = 0                                             # start with 1st vote history row in dictionary
    linesRead = 0                                       # No lines read as yet
    #
    #  Read Voter History File row by row ("VoterHistoryID","VoterID","Election Date","Vote Code")
    #  Accumulate for each voter into single record for configured elections only
    #
    for currentVoter in hDict["VoterID"]:
        linesRead += 1
        if (currentVoter != stateVoterID):
            #
            #  End of compilation for this voter, add calculated values to output record
            #
            if ( voterDataLine[0] == " " ):
                #
                #  Voter had no votes for the elections we're configured for
                #
                stateVoterID = currentVoter                             # Now processing this input record which is for the next voter
            else:
                #
                # Calculate rest of the values for this voter based on vote code
                #
                evaluateVoter()
                #
                # put caclulated values in output data line
                #
                voterDataLine[21] = str(votesTotal)
                voterDataLine[22] = str(generalCount)
                voterDataLine[23] = str(primaryCount)
                voterDataLine[24] = str(pollCount)
                voterDataLine[25] = str(absenteeCount)
                voterDataLine[26] = str(earlyCount)
                voterDataLine[27] = str(provisionalCount)
                voterDataLine[28] = str(voterRank)
                voterDataLine[29] = str(voterScore2)
                voterDataLine[0] = str(voterDataLine[0])                # make voter ID a string to write out
                temp = ",".join(voterDataLine) + "\n"
                stream.write(temp.encode('ASCII'))
                for z in VHdrCols:
                    voterDataLine[z] = ""                               # reset voter data acumulator for next voter
                linesWritten += 1
                linesIncWritten += 1
                totalVotes = 0
                if (linesIncWritten == 10000 ):
                    printLine(f"{linesWritten} lines written \r")
                    linesIncWritten = 0
                #
                stateVoterID = currentVoter                             # switch to next voter and process their first row
        #
        # for all records with this ID build out voterDataLine with all their in the 20 elections we're configured for 
        #
        voterDataLine[0] = currentVoter                                 # Set ID in output row in case 1st record for this voter
        votedate = hDict["Election Date"][row]                          # fetch election date from SOS vote history record
#        vdate = datetime.datetime.strptime(votedate,"%m/%d/%Y")         # make datetime object
        vdate = datetime.datetime(int(votedate[6:]), int(votedate[0:2]), int(votedate[3:5]), 0, 0 ,0)  # mm/dd/yyyy = yyyy,mm,dd
        if (vdate < oldestDate ):
            ignored += 1                                                # ignore records for elections older than we are looking for
            row += 1                                                    # fetch next row of SOS data
            continue
        #
        # find election this vote is for
        #
        flag = 0
        for cycle in CycleCols:
            if ((vdate >= voterEarlyDates[cycle]) and (vdate <= voterHeadingDates[cycle])):
                voterDataLine[cycle] = hDict["Vote Code"][row]          # This is the election this vote is for, stash it's code
                totalVotes += 1                                         # add to total votes for this voter
                flag = 1                                                # say election found
                break
        if (flag == 0):
            ignored += 1                                                # say this vote isn't for a configured election
        row = row+1
    #
    #  Done, close the output file
    #
    stream.seek(0)                                                      # rewind data stream
    hframe = pd.read_csv(stream)
    stream.close()
    printLine(f"<===> Completed Processing of: {voterHistoryFile}")
    printLine("<===> Total Vote History Records Read: {0:,}".format(linesRead))
    printLine("<===> Total Voter History Records Compiled: {0:,}".format(linesWritten))
    printLine("<===> Total Vote History Records for Non-configured Elections Ignored: {0:,}".format(ignored))
    #
    #  For lookup speed make nans = "" and make parallel lists of precincts, and dataframe rows
    #
    printLine(f"Transforming Voter History Array...")
    hframe = hframe.replace(np.nan, '', regex=True)                     # make any nans into '' for entire data frame
    hfvid = hframe['statevoterid'].tolist()                              # parallel list of voter IDs
    hflist = hframe.values.tolist()                                     # convert dataframe rows to python list of lists
    hframe= []                                                          # release dataframe memory
    #
    #   combined voter voting history records are now in dataframe "hframe"
    #
    #   Next read the Elegible Voter file and build combined base.csv record file
    #   and precinct breakout file.
    #
    printLine(f"Loading Eligible Voter File: {inputFile}. ")
    if (inputFile[-4:] == ".csv"):
        eframe = pd.read_csv (inputFile,low_memory=False)           #  Read .csv Eligible Voter file into dataframe "eframe"
    else:
        eframe = pd.read_excel (inputFile)                          #  Read .xls or .xlsx Eligible Voter file into dataframe "eframe"
    eframe = eframe.replace(np.nan, '', regex=True)                 #  make any nans into '' All data frame
    NumEv = len(eframe.index)
    #
    evHeadings = list(eframe.columns)                               # Get Vote History Headers
    for i in range(len(evHeadings)):
        evHeadings[i] = evHeadings[i].replace(" ", "")              # remove spaces from headings
        evHeadings[i] = evHeadings[i].replace("Residential", "")    # remove Residential if present
    printLine("{0:,} Eligible Voters Loaded...".format(NumEv))
    #
    #  Load County political district XREF file if present
    #
    if ( os.path.exists(adPoliticalFile) == False):
        printLine(f"******** Precinct XREF file {adPoliticalFile}  does not exist.")
        printLine("******** Output Base File Will Only Contain State Races, Local Races Will Be Blank!")
        Noxref = 1
    else:
        printLine(f"Loading Precinct XREF file {adPoliticalFile}")
        if (adPoliticalFile[-4:] == ".csv"):
            xframe = pd.read_csv (adPoliticalFile,low_memory=False) #  Read .csv Eligible Voter file into dataframe "eframe"
        else:
            xframe = pd.read_excel (adPoliticalFile)                #  Read .xls or .xlsx Eligible Voter file into dataframe "eframe"
        xframe = xframe.replace(np.nan, '', regex=True)             #  make any nans into '' All data frame
        xframe.sort_values(by=['PRECINCT'], inplace=True)           #  make sure it's sorted by precinct
        xreflist = xframe.to_dict(orient="list")                    #  get sorted xref DATAFRAME as Python Dictionary
        adPoliticalHeadings = list(xframe.columns)                  # Get XREF file Headers
    #
    #  Make dictionaries of input and output headings to allow indexing a list by name
    #
    bDict = {baseHeading[i]: i for i in range(len(baseHeading))} # dictionary of indexes to names for base.csv row
    eDict = {evHeadings[i]: i for i in range(len(evHeadings))}      # dictionary of indexes to Eligible Voter File row
    #
    printLine (f"Creating output base file {baseFile}")
    try:
        baseFileh = open(baseFile, "w")
    except IOError as e:
        printLine ("Unable tyo create base file: I/O error({0}): {1}".format(e.errno, e.strerror))
        exit(2)
    except: #handle other exceptions such as attribute errors
        printLine ("Unexpected error:", sys.exc_info()[0])
        exit(2)
    #
    temp = ",".join(baseHeading)                                # join header list into comma separated string
    print (temp, file=baseFileh)                                # Write out base.csv Header row


    # Open precinct.csv file
    printLine(f"Creating Voter precinct-table file: {pctFile}")
    try:
        pctFileh = open(pctFile, "w")
    except IOError as e:
        printLine ("Uanble to open Precinct file: I/O error({0}): {1}".format(e.errno, e.strerror))
        exit(2)
    except: #handle other exceptions such as attribute errors
        printLine ("Unexpected error:", sys.exc_info()[0])
        exit(2)
    temp = ",".join(pctHeading)                                # join header list into comma separated string
    print (temp, file=pctFileh)                                # Write out base.csv Header row\
    #
    # initialize the optional voter email log and the email array if selected
    #
    if (voterEmailFile != ""):
        printLine(f"Email updates file: {voterEmailFile}")
        try:
            emailLogFileh = open(emailLogFile, "w")
        except IOError as e:
            printLine ("Unable to open Email Update file: I/O error({0}): {1}".format(e.errno, e.strerror))
            exit(2)
        except: #handle other exceptions such as attribute errors
            printLine ("Unexpected error:", sys.exc_info()[0])
            exit(2)
        print (join(',',emailHeading),file=emailLogFileh)               # write header to email log file
        #
        if (voterEmailFile[-4:] == ".csv"):
            veframe = pd.read_csv (voterEmailFile,low_memory=False)     #  Read .csv Voter Email file into dataframe "eframe"
        else:
            veframe = pd.read_excel (voterEmailFile)                    #  Read .xls or .xlsx Voter Email file into dataframe "eframe"
        veframe = eframe.replace(np.nan, '', regex=True)                #  make any nans into '' All data frame
        #
        voterEmailHeadings = list(veframe.columns)                       # Get Voter email file headers
        #
        #  This get the email file into a dataframe, and the column header extracted to a list
        #  probably need to do more with this
        #
        # >>>>>>>>this has to be fixed for this to work <<<<<<<<<
    #
    # initialize the optional voter email log and the email array if selected
    #
    #
    #    Now iterate through the Eligible Voter frame rows and process each into the base file
    #
    BaseCol = len(baseHeading)
    linesIncRead = 0
    linesRead = 0
    for item in eframe.itertuples(name = None, index=False):
        #
        #  Read eligible Voter File row by row
        #
        line1Read = list(item)
        linesRead += 1
        linesIncRead += 1
        if ( linesIncRead > 4999 ):
            printLine(f"{linesRead} eligible voter records read\r")
            linesIncRead = 0
        #
        baseLine = [""] * BaseCol                                 # init baseLine to null entries
        #
        #  Copy across fixed values
        #
        voterid = line1Read[eDict["VoterID"]]                       # fetch voter ID of this entry
        baseLine[bDict["StateID"]] = makestr(voterid)
        baseLine[bDict["CountyID"]] = makestr(line1Read[eDict["CountyVoterID"]])
        baseLine[bDict["Status"]]   = makestr(line1Read[eDict["CountyStatus"]])
        baseLine[bDict["County"]]   = line1Read[eDict["County"]]
        baseLine[bDict["Precinct"]] = makestr(line1Read[eDict["RegisteredPrecinct"]])
        #
        #  Strip any CD, AD or SD down to only number if letters present
        #
        baseLine[bDict["CongDist"]] = makestr(line1Read[eDict["CongressionalDistrict"]]).replace("CD","")
        baseLine[bDict['AssmDist']] = makestr(line1Read[eDict["AssemblyDistrict"]]).replace("AD","")
        baseLine[bDict['SenDist']]  = makestr(line1Read[eDict["SenateDistrict"]]).replace("SD","")
        precinct = line1Read[eDict["RegisteredPrecinct"]]               # get this voter's precinct
        #
        #  get voting history record for this voter if it exists
        #
#        stats = list(hframe[hframe['statevoterid'] == currentVoter])
        ix = binarylookup(hfvid,voterid)
        if (ix >= 0):
            stats = hflist[ix]                                          # voterid found, fetch matching vote history record
            if (stats[0] != voterid):
                print (f"Lookup Error: ix={ix}, voterid = {voterid}")
                print (stats)
                exit(0)
            for i in range(0,29):
                baseLine[i+fixedflds] = makestr(stats[i+1])             # copy fields from voterdata record (all 20 cycles plus stats)
            statsAdded += 1
        else:
            # fill in record for registered voter with no vote history
            noData += 1                                                 # count voters with no vote history
            for i in CycleCols:
                baseLine[i+fixedflds] = ""                              # blank all 20 election votes
            baseLine[bDict["Generals"]]  = '0'
            baseLine[bDict["Primaries"]]    = '0'
            baseLine[bDict["Polls"]]        = '0'
            baseLine[bDict["Absentee"]]     = '0'
            baseLine[bDict["Early"]]        = '0'
            baseLine[bDict["Provisional"]]  = '0'
            baseLine[bDict["LikelytoVote"]] = "WEAK"
            baseLine[bDict["Score"]]        = '0'
            baseLine[bDict["TotalVotes"]]   = '0'
        if (baseLine[bDict["TotalVotes"]] == '0'):
            noVotes += 1                                                # count eligible voter with no votes
        #
        #  Fill In County Districts form Cross-reference file (if it exists)
        #
        if (Noxref == 0):
            ix = binarylookup(xreflist["PRECINCT"],precinct)                # find index of precinct in dictionary
            if (ix >= 0):
                #
                #  Found an XREF record for this precinct, Fill in the political districts from the XREF dictionary
                #
                baseLine[bDict["BrdofEd"]]    = makestr(xreflist["BOARDOFEDU"][ix])
                baseLine[bDict["CntyComm"]]   = makestr(xreflist["COMMISSION"][ix])
                baseLine[bDict["Rwards"]]     = makestr(xreflist["RWARDS"][ix])
                baseLine[bDict["Swards"]]     = makestr(xreflist["SWARDS"][ix])
                baseLine[bDict["SchBdTrust"]] = makestr(xreflist["SCHBDTRUST"][ix])
                baseLine[bDict["SchBdAtLrg"]] = makestr(xreflist["SCHBDATLRG"][ix])
            else:
                if (noPoliticalWarn == 0):
                    #
                    #  No XREF entry for this precinct
                    #
                    printLine ("******** WARNING!! YOU NEED TO UPDATE PRECINCT XREF FILE")
                    printLine (f"******** At least Precinct {precinct} not in precinct xref file.")
                    printLine ("******** File debug.txt lists all missing precincts.")
                    #
                    #  Open debug.txt to list missing precincts
                    #
                    debug = 1
                    try:
                        debugFileh = open(debugFile, 'w')
                    except IOError as e:
                        printLine(">>>>>I/O error({0}): {1}".format(e.errno, e.strerror))
                        debug = 0                                   # disable if for some reason doesn't open
                    except: #handle other exceptions such as attribute errors
                        printLine(">>>>>Unexpected error:", sys.exc_info()[0])
                        debug = 0                                   # disable if for some reason doesn't open
                    if (debug == 0):
                        printLine (">>>>> Could Not Create debug.txt file, proceeding without debug output")
                    noPoliticalWarn = 1                             # don't warn again of missing precincts on console
                if (debug != 0):
                    #
                    #  List all missing precincts in debug.txt, but not duplicates (same precinct missing in more than one voter record)
                    #
                    dup = 0
                    for i in range (0,len(duplicates)):
                        if ( precinct == duplicates[i]):
                            dup = 1                                       # already listed, skip listing it again
                            break
                    if (dup == 0):
                        #
                        #  List and remember a new missing precinct in debug.txt
                        #
                        duplicates.append(precinct)                        # add to duplicate missing precinct detection list
                        print(f"Precinct {precinct} not in precinct xref file", file= debugFileh)\
        #
        # convert proper names to upper case first then lower, then store in baseLine
        #
        phase = 1
        UCword = line1Read[eDict["FirstName"]].title()      # 1st letter UC, rest LC
        baseLine[bDict["First"]] = UCword
        ccfirstName = UCword                                # Save first name for email lookup

        phase = 2
        baseLine[bDict["Middle"]] = line1Read[eDict["MiddleName"]].title()    # 1st letter UC, rest LC

        phase = 3
        UCword = line1Read[eDict["LastName"]].title()       # 1st letter UC, rest LC
        UCword = UCword.replace(" ","")                     # remove all imbedded spaces
        UCword = UCword.replace(",","-")                    # change comma to dash
        baseLine[bDict["Last"]] = UCword
        cclastName = UCword                                 # save last name for email lookup
        #
        #  Copy Rest of SOS registration file fields to baseLine
        #
        baseLine[bDict["BirthDate"]]  = line1Read[eDict["BirthDate"]]
        baseLine[bDict["RegDate"]]    = line1Read[eDict["RegistrationDate"]]
        baseLine[bDict["Party"]]      = line1Read[eDict["Party"]]
        baseLine[bDict["Phone"]]      = line1Read[eDict["Phone"]]
        UCword                        = line1Read[eDict["Address1"]].title()
        baseLine[bDict["Address1"]]   = UCword
        if (UCword == ""):
            baseLine[bDict["StreetNo"]]   = ""                              # Address1 was empty, so both parts are too
            baseLine[bDict["StreetName"]] = ""
        else:
            streetno = UCword.split(" ",1)                                  # split street number from street name
            baseLine[bDict["StreetNo"]]   = streetno[0]
            baseLine[bDict["StreetName"]] = streetno[1]
        baseLine[bDict["Address2"]]   = line1Read[eDict["Address2"]].title()
        baseLine[bDict["City"]]       = line1Read[eDict["City"]].title()
        baseLine[bDict["State"]]      = line1Read[eDict["State"]]
        baseLine[bDict["Zip"]]        = str(line1Read[eDict["Zip"]])
        baseLine[bDict["email"]]      = ""                                  # default to no email at this point
        #
        #  do any email matching if selected on command line
        #
        #   >>>>>>>>>>>  THIS NEEDS TO BE TESTED AND PROBABLY FIXED  <<<<<<<<<<<<<<<<
        #
        if (voterEmailFile != ""):
            #  locate email address if available
            #  "Last", "First", "Middle","Phone","email","Address", "City","Contact Points",
            #     0       1         2                4      5          6          7
            emails = veframe.loc[(veframe['Last'] == cclastName) & (veframe['First'] == ccfirstName)]
            if (emails.empty == False):
                voterEmailArray = list(emails)
                calastName               = voterEmailArray[0]
                cafirstName              = voterEmailArray[1]
                caemail                  = voterEmailArray[4]
                baseLine[bDict["email"]] = voterEmailArray[4]
                capoints                 = voterEmailArray[7].replace(";",",")
                emailAdded = emailAdded + 1

                # build a trace line to show email was updated
                emailLine = ["","","","",""]
                emailLine[0] = str(voterid)
                emailLine[1] = str(line1Read[eDict["RegisteredPrecinct"]])
                emailLine[2] = calastName
                emailLine[3] = cafirstName
                emailLine[4] = caemail
                emailProfile = []
                try:
                    for i in range(0,5):
                        emailProfile.append(emailLine[i])
                    print (join( ',', emailProfile ), file=emailLogFileh)
                except TypeError:
                    print ("error:", sys.exc_info()[0])
                    print (emailLine)
                    exit(2)
        #
        #-------------------------------------------------
        # caclulate registered days and age of voter
        #
        birthdate = line1Read[eDict["BirthDate"]]
        regdate   = line1Read[eDict["RegistrationDate"]]
        before = 0

        # determine age
        # my ( date, yy, mm, dd, now, age, regdays, before, adjustedDate );
        if (birthdate != ""):
            #before = datetime.datetime.strptime(birthdate,"%m/%d/%Y")
            before = datetime.datetime(int(birthdate[6:]),  int(birthdate[0:2]), int(birthdate[3:5]), 0, 0, 0)
            now          = datetime.datetime.today()
            age          = now - before
            age          = age.days / 365                   # age in years
            age          = str(round(age))                  # get integer of age in years
        else:
            age = ""                                        # birthday nt present
        baseLine[bDict["Age"]] = age                        # store age in base record
        #
        # determine registered days
        #
        #before = datetime.datetime.strptime(regdate,"%m/%d/%Y")
        before = datetime.datetime(int(regdate[6:]), int(regdate[0:2]), int(regdate[3:5]),0 ,0 ,0)
        regTimePiece = before;                                            # save encoded registration date for later work
        now          = datetime.datetime.today()
        regdays      = now - before
        regdays      = regdays.days
        baseLine[bDict["RegisteredDays"]] = str(regdays)
        #
        #-------------------------------------------------------------------------------
        #  Find oldest election reg date allows vote in.  If older vote, use that date
        #  as it means voter re-registered at some point.
        #
        #  Calculate propensity strength from possible votes vs. actual votes
        #
        rstop = 0
        ovote=""
        test = 0
        vid = line1Read[eDict["VoterID"]]
        for j in range(0, 20):
            edate = voterHeadingDates[j+1]
            if (edate > regTimePiece):
                # voter was registered for this election
                rstop = j;                                                    # index+1 to oldest election registered for
            else:
                #
                #  See if older vote than registration date
                #
                if (baseLine[j+fixedflds] != ""):
                    rstop = j;                                                # must have re-registered
                    test = 1
        #
        #  rstop = index to oldest possible vote for this voter.
        #  calculate voter propensity to vote strength based on
        #  this many possible votes.
        #
        maxstrength = 0
        voterstrength = 0                      # init accumulators
        c = 0
        ev = 0
        for j in range(0,rstop+1):
            maxstrength = maxstrength + electionValue[j]                      # sum possible election strengths
            if (baseLine[j+fixedflds] != ""):
                voterstrength = voterstrength + electionValue[j]              # sum actual voted election strengths
        voterstrength = ((voterstrength/maxstrength) * 10)                    # calc voter strength 0-9.99
        baseLine[bDict["Score"]] = str(round(voterstrength))
        if (voterstrength <= 2):
            baseLine[bDict["LikelytoVote"]] = "WEAK"                          # < 2 = weak
        if (voterstrength == 0):
            baseLine[bDict["LikelytoVote"]] = "NEVER"                         # special case strength of 0
        if (int(baseLine[bDict["Primaries"]]) > 0):
            baseLine[bDict["LikelytoVote"]] = "MODERATE"                      # Moderate if voted in a primary
        if ((voterstrength > 2) and (voterstrength <= 6)):
            baseLine[bDict["LikelytoVote"]] = "MODERATE"
        if (voterstrength > 6):
            baseLine[bDict["LikelytoVote"]] = "STRONG"
        voterstrength = int(voterstrength + 0.49)                             # convert to 0-10 score
        #
        #   base.csv record complete, write it out
        #
        #
        i = 0
        for item in baseLine:
            if "," in item:
                baseLine[i] = "\"" + baseLine[i] + "\""
            i = i + 1
        #
        try:
            print (",".join(baseLine), file = baseFileh)            # write this voter's record to output file
        except TypeError:
            print ("error:", sys.exc_info()[0])
            print (baseLine)
            exit(2)
        #
        #  Do the precinct stats accumulation for this record
        #
        calc_precinct()
    #
    #   Done, print summary and exit
    #
    printLine(f"Writing precinct summary file {pctFile}...")
    print (f"# precints in list = {len(PctPrecinct)}")
    write_precinct()                                            # write the precinct.csv file
    pctFileh.close()                                            # close precinct file
    baseFileh.close()                                           # close output file
    if (voterEmailFile != "" ):
        emailLogFileh.close()
    printLine("<===> Total Eligible Voter Records Read: {0:,}".format(linesRead))
    printLine("<===> Total Voting History Stats added: {0:,}".format(statsAdded))
    printLine("<===> Total Registered Voters with no Recent Vote History: {0:,}".format(noVotes))
    printLine("<===> Total Registered Voters with no Vote Record: {0:,}".format(noData))
    printLine("<===> Total Email Addresses added: {0:,}".format(emailAdded))
    printLine("<===> Total Precincts found and {0} Records written: {1:,}".format(pctFile, NumPct))
    printLine("<===> Total base.csv Records written: {0:,}".format(linesWritten))
    #
    #   Program is done, Print Overall Run Time, close log file and exit
    #
    EndTime = time.time()
    TotSec = int((EndTime - StartTime)*10)/10
    TotMin = int (TotSec/60)
    if (TotMin > 0):
        TotSec = int((TotSec -(TotMin*60))*10)/10
        printLine (f"Total Elapsed time is {TotMin} Minutes {TotSec} seconds.\n")
    else:
        printLine (f"Total Elapsed time is {TotSec} seconds.\n")
    printFileh.close()
    return (0)
    #
    #  End of program
    #
    

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
