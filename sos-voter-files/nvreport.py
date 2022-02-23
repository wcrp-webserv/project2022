#************************************************************************************
#                                 nvreport.py                                       *
#                                                                                   *
#  Input is one or more i360 survey file(s) and optional question text file         *
#        may be in either .csv or .xlsx format                                      *
#                                                                                   *
#  Output is a text report of the survey and analysis thereof  jes                  *
#                                                                                   *
# *********************************************************************************** 

import pandas as pd
import numpy as np
import sys, getopt, os
from os import listdir
import shutil, fnmatch
from os.path import isfile, join
import xlsxwriter
import calendar
from datetime import datetime
import time
import math

#==================================================================================================================
#
#               +-----------------------------------------------+
#     > > > > > | D E F I N E   G L O B A L   V A R I A B L E S | < < < < <
#               +-----------------------------------------------+
#
###################################################################################################################
# >>>>>>>>>>>>>>>>>>>>   Here is where Fixed Assumptions About SpreadSheet are Defined   <<<<<<<<<<<<<<<<<<<<<<<< #
# >>>>>>>>>>>>>>>>>>>>           Change these if Spreadsheet format Changes!!!!          <<<<<<<<<<<<<<<<<<<<<<<< #
#                                                                                                                 #
#  Response Text Strings with special meaning                                                                     #
#                                                                                                                 #
Take = "Take Survey"        # This Response says this person TOOK Survey                                          #
Refuse = "Refused"          # This Response Says person answered but REFUSED TO TAKE Survey                       #
#                                                                                                                 #
#  Definition of Fixed Columns which allow locating the Question Columns and certain fixed data items             #
#       Note: Column A in Excel = Column 0 here, Column 1 = B etc.                                                #
#                                                                                                                 #
VolFirstText = "Volunteer First Name"           # Column Heading for Volunteer First Name                         #
VolLastText = "Volunteer Last Name"             # Column Heading for Volunteer Last Name                          #
PhoneText = "Phone"                             # Column Heading of Phone Number Called                           #
ResponseText = "Household response"             # Column Heading for Contact Response (Busy, No Answer, etc.)     #
DateText = "Response Date"                      # Column Heading for Contact Date                                 #
FirstQCol = 15                                  # Column of First Question in Row = 15 (col P in Excel)           #
ColsAfterQ = 3                                  # Number of Columns Following Questions = 3                       #
#                                                                                                                 #
#  These exact text responses allow sorting answers and also allow detection of a survey respondent's self        #
#  definition of their political persuasion.                                                                      #
#                                                                                                                 #
DefAns1 = ("Strongly Disagree","Disagree","No Opinion","Agree","Strongly Agree")                                  #
DefAns2 = ("Conservative","Moderate Conservative","Moderate","Moderate Progressive","Progressive")                #
#                                                                                                                 #
# >>>>>>>>>>>>>>>>>>>>>>>>>    END OF FIXED ASSUMPTIONS ABOUT SPREADSHEET FORMAT    <<<<<<<<<<<<<<<<<<<<<<<<<<<<< #
###################################################################################################################

svyfile  = ""                                                                   #  Input: survey file ( -infile)
svyfile2  = ""                                                                   #  Input: survey file ( -infile)
svydir = ""                                                                     #  Input: directory of survey files (-survey)
qfile = ""                                                                      #  Input: question text file (-qfile)
datadir = ""                                                                     # Output: Survey Report directory (-datadir)
rptfile = ""                                                                    # Output: Survey Report file (-rptfile)
outfh = ""                                                                      # Output: Survey Report file handle

printFile = "print.txt"                                                         # Console Log File
printFileh = ""                                                                 # console Log file handle (must be global)
#
#  -Survey directory search, -select and -pers persuasion qualifier variables
#
PersOpts =[]                   #Persuasion qualifier Option Array
NumPersOpts = 0                # Number of persuasion qualifier options
SelPhase = []                  # phase selection params
SelDist = []                   # district selection params
SelParty = []                  # party selection params
SelLikely = []                 # voting propensity selection params
SurveyFiles = []               # list of 1 to n survey files to report
#
#   Survey SpreadSheet Information (from init_spreadsheet)
#
Headings =[]                   # Array of Text Headings for spreadsheet
NumSvyRows = 0                 # Number of Rows in current Spreadsheet
NumSvyCols = 0                 # Number of Columns in current Spreadsheet
SurveyTitle =""                # Title of this survey (from Column 1 or spreadhseet)
DateCol = -1                   # Call Date Column Index
VolFirstCol = -1               # Volunteer First Name Column Index
VolLastCol = -1                # Volunteer Last Name Column Index
PhoneCol = -1                  # Phone Number Called Column Index
ResponseCol = -1               # Response Code for Call/Contact Column Index
#
#  Response Types And Counts
#
#  Numresponse is Number of unique responses
#  ResponseText = Array of text responses
#  ResponseCnt = parallel array of Number of time this text response was given
#
NumResponse = 0                # Number of Response Types
ResponseCode =[]               # Array of reponse texts
ResponseCnt =[]                # Array of Times Each Text Logged
#
#  Volunteer Names and Call Attempt & Result Counts
#
NumVolunteer = 0               # of different Volunteer Names
VolunteerName =[]              # Array of Volunteer Names     
VolunteerAttempts =[]          # Array of # of call attempts by each volunteer
VolunteerAnswers =[]           # Array of # of answers by Volunteer
VolunteerContacts =[]          # Array of number of Surveys Taken by Volunteer
VolunteerRefuse =[]            # Array of survey refusals by Volunteer
VolunteerPartial =[]           # Array of # of Partial Surveys by volunteer (person quit partway thru)
#
#   Totals for All Volunteers Combined
#
NumAttempts = 0                # Total Attempts for all volunteers
NumAnswers = 0                 # Total Answers for all volunteers
NumContacts = 0                # Total Surveys Taken for all volunteers
NumRefuse = 0                  # Total Refusals for all volunteers
NumPartial =0                  # Partial Surveys
#
#  Phone Number Tables to find Unique Households Called
#
NumPhones = 0                  # Total Unique Phone Numbers
PhoneNumber =[]                # List of Phone Numbers Dialed
PhoneDuplicates = 0            # Number of times redialed any number
#
# By Day Of Week Tables
#
DowIndex = 0;                   # Day of Week Index (0 to 6)
DowText = ("Sun","Mon","Tue","Wed","Thu","Fri","Sat")
DowAttempts = [0,0,0,0,0,0,0] # Unique Call Attempts by Day Of Week
DowAnswers = [0,0,0,0,0,0,0]   # Answers by Day Of Week
DowContacts = [0,0,0,0,0,0,0]  # Surveys Taken by Day of Week
DowRefuse = [0,0,0,0,0,0,0]    # Surveys Refused by Day of Week
#
#  Question Global Tables
#
NumQuestions = 0               # Number of Questions in this spreadsheet (Discovered from SpreadSheet Header Processing)
QuestionText = []              # Table of Question Text (Copied from SpreadSheet Header Row)
QuestionNumResponse = []       # Parallel Table of # of responses for each question
QuestionResponse = []          # Parallel Array of Pointers to Table of Response Text for each question
QuestionTally = []             # Parallel Array of Pointers to Table of Count for each Response Text for each Question
QuestionDecline = []           # Parallel Array of pointers to Table of count for those who declined to answer this question
TotQuestions = 0               # total questions across all spreadsheets
QuestionIndex =[]              # list of indexes into Parallel tables for this spreadsheet
#
sframe=""                                                                       # Survey File DataFrame once loaded
#
ProgName = ""                  # Name of running progrAM (from sys.arv[0])
#               +-----------------------------------------------+
#     > > > > > | E N D   O F   G L O B A L   V A R I A B L E S | < < < < <
#               +-----------------------------------------------+
#=====================================================================================================================
#
#************************************
#                                   *
# Print Log line to screen and file *
#                                   *
#************************************
def printLine (printData):
    global printFileh,ProgName
    datestring = datetime.now()
    datestring = datestring.strftime("%m/%d/%Y %H:%M:%S")
    print( ProgName + " " + datestring + ' - ' + printData)
    print( ProgName + " " + datestring + ' - ' + printData, file=printFileh)
    return

def printhelp():
    print ("py nvreport.py -infile <filename> -datadir <directory> -rptfile<filename> -survey <path> -select param,param,...")
    print ("    -infile <filename> reports from a single file. Overrides -survey option if both present.\n")
    print ("    -survey <path> reports from survey files in the specified directory. ")
    print ("           Note: If -infile=i360 then the current directory  ")
    print ("                 for a name of SurveyResponse* and that file will be renamed to the name   ")
    print ("                 of the survey specified in the column 'Title' and the   ")
    print ("                 file is moved to the location specified by <-datadir>/surveys/ ")
    print ("           Note: In the absence of either -infile or -survey the current working directory")
    print ("           will be used as if a -survey <cwd> were specified.\n")
    print ("    -select specifies which files in the survey directory will be selected. ")
    print ("           ADnn - selects files from matching assembly district.")
    print ("           SDnn - selects files from matching senate district.")
    print ("             pn - selects files from the specified survey phase.")
    print ("            rep - selects file that surveyed republican voters.")
    print ("            dem - selects file that surveyed democrat voters.")
    print ("            oth or othr - selects file that surveyed other party voters.")
    print ("           high - selects file that surveyed high propensity voters.")
    print ("            mod - selects file that surveyed moderate propensity voters.")
    print ("            low - selects file that surveyed low propensity voters.")
    print ("              Note: parameters can be combined in any way. Example:")
    print ("                    -select p0,AD27,rep,oth,high,mod\n")
    print ("    -pers specifies that only some respondent persuasion(s) will be compiled.")
    print ("              C - selects voters who self select as Conservative.")
    print ("              MC - selects voters who self select as Moderately Conservative.")
    print ("              M - selects voters who self select as Moderate.")
    print ("              MP - selects voters who self select as Moderately Progressive.")
    print ("              P - selects voters who self select as Progressive.")
    print ("              Note: parameters can be combined in any order. Example:")
    print ("                    -pers C,M,MC\n")
    print ("    -qfile <filename> text file of question text to substitute for the I360 text.\n")
    print ("    -datadir <directory> specifies the output report directory.  Default is null")
    print ("    -rptfile <filename> specifies the output report file.  Default is report.txt")
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
    global datadir, rptfile, svydir, svyfile, qfile                          # File/directory parameters
    global NumPersOpts, PersOpts, DefAns2                           # persuation parameters
    global SelPhase, SelDist, SelParty, SelLikely                   # selection parameters

    selstring=""                                                    # Assume no selection options
    perstring =""                                                   # assume no persuasion options
    print("")
    NumParms = len(argv)                                            # get number of items in argument list
    if (NumParms == 0):
        printhelp()                                                 # No args, print help and exit
        exit(2)
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
        elif opt == "-datadir":
            datadir = arg
        elif opt == "-rptfile":
            rptfile = arg
        elif opt == "-infile":
            svyfile = arg           
            if (svyfile == 'i360'):
                Dir = os.getcwd()
                #svyfile = os.path.join(Dir)
                #svyfile = svyfile + "/SurveyResponse*"       # form fully qualified survey file name
                for sfile in os.listdir('.'):
                    if fnmatch.fnmatch(sfile, 'SurveyResponse*'):
                        svyfile = sfile
                        print ( svyfile + '\n')
                        break
                if (os.path.isfile(svyfile) == False):
                    printLine (f"Survey File '{svyfile}' does not exist...aborting\n")
                    exit(2)
            elif(svyfile[2] != ":"):
                Dir = os.getcwd()                                   # get our current working directory
                svyfile = os.path.join(Dir,svyfile)                 # form fully qualified survey file name
                if (os.path.isfile(svyfile) == False):
                    printLine (f"Survey File '{svyfile}' does not exist...aborting\n")
                    exit(2)

        elif opt == "-survey":
            svydir = arg
        elif opt == "-qfile":
            qfile = arg
        elif opt == "-select":
            selstring = arg
        elif opt == "-pers":
            perstring = arg
        else:
            print (f">>>UnKnown Option: {opt}\n")
            printhelp()                                     # Unknown Option - provide help and exit
            exit(2)
    #
    #  Arguments gathered, now process any sub-params given
    #
    #----------------------------------------------------------
    #  Handle -select parameters if specified on command line
    #----------------------------------------------------------
    if (selstring != ""):
        #
        #  List the options we scanned
        #
        opts = selstring.split(",")
        temp = "-select options are: "
        #
        #  Now stash them into the arrays by their type
        #
        for i in range(0,len(opts)):
            temp = f"{temp}\"{opts[i]}\" "                                  # add to list of options
            chr = opts[i][0].lower()                                        # get 1st character  of optin lower case
            if (chr == "p" ):
                SelPhase.append(opts[i].lower())                            # stash phase selector
            elif ((chr == "a") or (chr == "s")):
                SelDist.append(opts[i].lower())                             # stash district selector
            elif ((chr == "r") or (chr == "o") or (chr == "d")):
                if (opts[i].lower() == "othr"):
                    opts[i] = "oth"                                         # make "oth" and "othr" synonyms
                SelParty.append(opts[i].lower())                            # stash party selector
            elif ((chr == "h") or (chr == "m") or (chr == "l")):
                SelLikely.append(opts[i].lower())                           # stash propensity selector;
            else:
                #
                #  We don't recognize this option type
                #
                printLine(f"Invalid -select parameter \"{opts[i]}\" ignored! \n")
        printLine(f"{temp}")

    #----------------------------------------------------------
    #  Handle -pers parameters if specified on command line
    #----------------------------------------------------------
    if (perstring != ""):
        #
        #  List the options we scanned
        #
        opts = perstring.split(",")
        temp = "-pers restrictions are: "
        #
        #  Now stash them into the arrays by their type
        #
        for i in range(0,len(opts)):
            chr = opts[i].upper()
            if (chr == "C" ):
                PersOpts.append(DefAns2[0])                             # stash persausion selector
                NumPersOpts += 1
                temp = f"{temp}\"{DefAns2[0]}\" "
            elif (chr == "MC" ):
                PersOpts.append(DefAns2[1])                             # stash persausion selector
                NumPersOpts += 1
                temp = f"{temp}\"{DefAns2[1]}\" "
            elif (chr == "M" ):
                PersOpts.append(DefAns2[2])                             # stash persausion selector
                NumPersOpts += 1
                temp = f"{temp}\"{DefAns2[2]}\" "
            elif (chr == "MP" ):
                PersOpts.append(DefAns2[3])                             # stash persausion selector
                NumPersOpts += 1
                temp = f"{temp}\"{DefAns2[3]}\" "
            elif (chr == "P" ):
                PersOpts.append(DefAns2[4])                             # stash persausion selector
                NumPersOpts += 1
                temp = f"{temp}\"{DefAns2[4]}\" "
            else:
                #
                #  We don't recognize this persuasion type
                #
                printLine(f"Invalid -pers parameter \"{opts[i]}\" ignored! \n")
        printLine(f"{temp}")
    
    if (rptfile == ""):
        rptfile = "report.txt"
    return
#
#=====================================================================================================
#
#--------------------------------------------------------------
#   Initialize key items for newly opened spreadsheet header
#--------------------------------------------------------------
#
#  1. Find size of spreadsheet (#rows and # cols)
#  2. Read header row into @Headings array
#  3. Find what columns are the Volunteer name (first & last)
#  4. Find what columns are the phone number and contact response
#  5. Find what column is the contact attempt date
#  6. Verify that these things were found
#
#  Returns 0 if init succeeded  error code 1-5 if failure
#
def init_spreadsheet():
    global VolFirstText,  VolLastText, PhoneText, ResponseText, DateText                # Configuration Inputs
    global Headings, NumSvyRows, NumSvyCols, SurveyTitle, DateCol                       # Output fields
    global VolFirstCol, VolLastCol, PhoneCol, ResponseCol                               #   ...

    NumSvyRows = len(sframe.index)                      # get Number of Rows in Spreadsheet
    NumSvyCols = len(sframe.columns)                     # get Number of Columns in Spreadsheet
    #
    # Fetch and save Excel Header row text strings into Headings list
    #
    Headings = list(sframe.columns)                     # Get Column Names in List
    #
    #  Locate the Columns we will need to know by their Column Header Text
    #
    nfnd = 0
    for j in range(0,len(Headings)):
        Colhdr = Headings[j].lower()                    # Get header of This Column in lower case
        if (Colhdr == VolFirstText.lower()):
            VolFirstCol = j                             # Save Column index for Volunteer First Name
        elif (Colhdr == VolLastText.lower()):
            VolLastCol = j                              # Save Column index for Volunteer Last Name
        elif (Colhdr == PhoneText.lower()):
            PhoneCol = j                                # Save Column index for Phone Number Called
        elif (Colhdr == ResponseText.lower()):
            ResponseCol = j;                            # Save Column index for Call Response
        elif (Colhdr == DateText.lower()):
            DateCol = j                                 # Save Column index for Call Date
    #
    #  Be Sure We Found the columns Headings We Need
    #
    if (VolFirstCol == -1):
        printLine(f"Couldn't Find Required Column Headed {VolFirstText} ...")
        return(1)
    if (VolLastCol == -1):
        printLine(f"Couldn't Find Required Column Headed {VolLastText} ...")
        return(2)
    if (PhoneCol == -1):
        printLine(f"Couldn't Find Required Column Headed {PhoneText} ...")
        return (3)
    if (ResponseCol == -1):
        printLine(f"Couldn't Find Required Column Headed {ResponseText} ...")
        return (4)
    if (DateCol == -1):
        printLine(f"Couldn't Find Required Column Headed {DateText}...")
        return (5)
    #
    #  Get Title Text for this survey
    #
    SurveyTitle = sframe.iloc[0][0]
    printLine (f"... Survey Title: {SurveyTitle}")
    #
    return 0
#
#==================================================================================================
#
#--------------------------------------------------------------------
# see if this file in the -survey directory qualifies to be selected
#
#  title = Survey title and has one of two formats:
#
#       i-20p<p>-<dist>-<party>-<likely>-...
#   or  
#       p<p>-<dist>-<party>-<likely>-...
#
#       p = phase of this survey
#       dist = ADn or SDn district of this survey
#       party = party of the respondents to this survey
#       likely = the voting propensity of the respondents
#
#  SelPhase = list of phase command line -select criteria
#  SelDist = list of district command line -select criteria
#  SelParty = list of party command line -select criteria
#  SelLikely = list of voting propensity command line -pers criteria
#  
#--------------------------------------------------------------------
def qualify_file(title):
    #
    #  Selection was requested, Scan survey title for option params
    #
    temp = title.lower()                                            # force title lower case
    opts = temp.split("-")                                          # break out lower cased title into segments
    if (opts[0] == "i"):
        leadI = 1                                                   # it is i-20pn format, add 1 to indices
        temp = opts[1][2:]                                          # and extract pn param from 20pn format
    else:
        leadI = 0                                                   # it is pn- format, add 0 to indices
        temp = opts[0]                                              # get pn param
    phase = temp;                                                   # store phase
    dist = opts[1+leadI]                                            # store district
    party = opts[2+leadI]                                           # store party
    if (party == "othr"):
        party = "oth"                                               # make "othr" match "oth"
    likely = opts[3+leadI]                                          # store propensity
    #
    #  Params broken out of title, now check if title qualifies for selection or not
    #
    if(len(SelPhase)+len(SelDist)+len(SelParty)+len(SelLikely) == 0):
        return 0                                                    # no selection criteria, file qualifies
    t = -1                                                          # assume failure to select  
    x = 0
    if (len(SelPhase) > 0):
        for x in range(0, len(SelPhase)):                           # phase parameter specified
            if (SelPhase[x] == phase):
                t=0;                                                # file Phase is selected
                break;                                              # move to any other params
        if (t == -1):
            return -1                                               # file not selected
    t = -1
    if (len(SelDist) > 0):                                          # district parameter specified
        for x in range(0, len(SelDist)):
            if (SelDist[x] == dist):
                t=0                                                 # district is selected
                break                                               # move to any other params
        if (t == -1):
            return -1                                               # file not selected
    t = -1
    if (len(SelParty) > 0):                                         # party parameter specified
        for x in range(0, len(SelParty)):
            if (SelParty[x] == party):
                t=0                                                 # party is selected
                break                                               # move to any otehr params
        if (t == -1):
            return -1                                               # file not selected
    t = -1
    if (len(SelLikely) > 0):                                        # propensity parameter specified
        for x in range (0, len(SelLikely)):
            if (SelLikely[x] == likely):
                t=0                                                 # propensity is selected
                break                                               # move to any otehr params
        if (t == -1):
            return -1                                               # file not selected
    return 0;                                                       # file is selected
#
#====================================================================================================
#
#-----------------------------------------------------------------
#            Process Questions for this row
#
#   rlist = list of columns in this row
#
#   return 0 if all questions answered, 1 if only some answered
#-----------------------------------------------------------------
def doquestions (rlist):
    global NumQuestions, QuestionNumResponse, QuestionResponse, QuestionTally, QuestionDecline
    #
    rcode = 0                                           # assume all questions answered
    #
    #  cycle thor0ugh the questions and log the responses.  Build response list for each question as we go
    #
    for q in range(0,NumQuestions):
        rtext = (rlist[FirstQCol+q])                    # fetch response to this question
        if (pd.isna(rtext)):
            rtext = ""                                  # Handle Pandas NAN on .xlsx load of empty cell
        if (rtext == ""):
            QuestionDecline[q] += 1                     # add declined to answer
            rcode = 1                                   # flag Partial Survey
            continue                                    # skip this question
        z = QuestionIndex[q]                            # get the question index into the parallel table array
        if (QuestionNumResponse[z] == 0):
            QuestionResponse[z].append(rtext)           # First time Response for This question
            QuestionTally[z].append(1)                  # It has Happened once
            QuestionNumResponse[z] = 1                  # this question now has 1 response
        else:
            try:
                i = QuestionResponse[z].index(rtext)    # see if this respinse already given for this question
                QuestionTally[z][i] += 1                # Another hit for this response, count it
            except ValueError:
                QuestionResponse[z].append(rtext)       # New Response for This question, add to response list
                QuestionTally[z].append(1)              # It has Happened once
                QuestionNumResponse[z] += 1             # this question now has another response
    return rcode                                        # all questions done, return full or partial answers
#
#====================================================================================================
#
#----------------------------------------------------------------
#   Generate report header plus the Call Attempt/Result Report
#----------------------------------------------------------------
def ReportAttempts(pfirst,plast):
    global NumAttempts, NumContacts, NumPartial, PhoneDuplicates, NumRefuse, NumAnswers, outfh

    printLine(f"Survey Covers {NumAttempts} Calls with {NumAnswers} Answers")
    #
    #  Calculate and/or format data items
    #
    attempts = "{0:4d}".format(NumAttempts)
    Contacts = "{0:4d}".format(NumContacts)
    partial = "{0:4d}".format(NumPartial)
    surveyPct = "{0:4.1f}".format(NumContacts/NumAttempts*100)
    redial =  "{0:4d}".format(PhoneDuplicates)
    refuses = "{0:4d}".format(NumRefuse)
    refusePct = "{0:4.1f}".format(NumRefuse/NumAnswers*100)
    unique = "{0:4d}".format(NumAttempts-PhoneDuplicates)
    Answers = "{0:4d}".format(NumAnswers)
    AnswerPct = "{0:4.1f}".format(NumAnswers/NumAttempts*100)
    #
    # Print the report lines
    #
    if (len(SurveyFiles) == 1):
        print (f"                 Reporting From Survey {SurveyTitle}", file=outfh)
    else:
        print ("                 Reporting From Combined Files:", file = outfh)
        for i in range(0,len(SurveyFiles)):
            print (f"                 {SurveyFiles[i]}", file=outfh)
    if (NumPersOpts > 0):
        #
        #  print that  this is conditional partial report of selected persuasion respondents
        #
        print (f"\n   Reporting only Persuasion(s) {','.join(PersOpts)}\n", file = outfh)
    print (f"\n\n                 Survey Interval from {pfirst} to {plast}\n\n", file=outfh)
    print ("                      Total Survey Call Effectiveness", file=outfh)
    print ("                      -------------------------------", file=outfh)
    print (f"Call Attempts: {attempts}       Total Answers: {Answers}               Answer%: {AnswerPct}%", file=outfh)
    print (f"      Redials: {redial}     Survey Contacts: {Contacts} (Part {partial})   Survey%: {surveyPct}%", file=outfh)
    print (f" Total Unique: {unique}     Survey Refusals: {refuses}               Refuse%: {refusePct}%\n\n", file=outfh)
    return
#
#====================================================================================================
#
#---------------------------------------------------------------
#  Generate Call Response Summary Report
#---------------------------------------------------------------
def ReportCallStatus():
    global NumResponse, ResponseCode, outfh

    printLine(f"Total Call Response Types: {NumResponse}")
    print ("                         Incomplete Call Breakdown", file=outfh)
    print ("                         -------------------------", file=outfh)
    for i in range(0,NumResponse - 1):
        if (ResponseCode[i] == Take):
            continue
        if (ResponseCode[i] == Refuse):
            continue
        print (f"{ResponseCode[i]} {ResponseCnt[i]}, ", end="", file=outfh)
    last = NumResponse-1
    print (f"{ResponseCode[last]} {ResponseCnt[last]}\n\n", file=outfh)
    return
#
#====================================================================================================
#
#---------------------------------------------------------------
#  Generate the by Volunteer Report
#---------------------------------------------------------------
def ReportByVolunteer():
    global NumVolunteer, VolunteerName, VolunteerAttempts, VolunteerAnswers, VolunteerContacts
    global VolunteerRefuse, VolunteerPartial, outfh

    printLine(f"Total Volunteer Names: {NumVolunteer}")
    name = ""                  # Define local working variables
    attempts = ""
    Answers = ""
    Contacts = ""
    refuse = ""
    partial = ""
    completePct = ""
    AnswerPct = ""
    print ("                         Performance by Volunteer", file=outfh)
    print ("                         ------------------------\n", file=outfh)
    print ("     Volunteer Name   Attempts  Answers   Contacts  Refuse  Completion%  Answer%", file=outfh)
    print ("     --------------   -------   -------   --------  ------  -----------  --------", file=outfh)
    for i in range(0,NumVolunteer):
        name = "{0:>19s}".format(VolunteerName[i])                  # format items to print
        attempts = "{0:4d}".format(VolunteerAttempts[i])
        Answers = "{0:4d}".format(VolunteerAnswers[i])
        Contacts = "{0:4d}".format(VolunteerContacts[i])
        refuse = "{0:4d}".format(VolunteerRefuse[i])
        partial= "{0:2d}".format(VolunteerPartial[i])
        completePct = "{0:4.1f}".format(VolunteerContacts[i]/VolunteerAttempts[i]*100)
        AnswerPct = "{0:4.1f}".format(VolunteerAnswers[i]/VolunteerAttempts[i]*100)
        if (VolunteerPartial[i] != 0):
            print (f"{name}     {attempts}      {Answers}    {Contacts}({partial})   {refuse}      {completePct}%       {AnswerPct}%", file=outfh)
        else:
            print (f"{name}     {attempts}      {Answers}    {Contacts}       {refuse}      {completePct}%       {AnswerPct}%", file=outfh)
    print ("\n", file=outfh)
    return
#
#====================================================================================================
#
#--------------------------------------------------------------
#  Generate the Question & Answer report
#--------------------------------------------------------------
#
#  Two answer sets that if we find them, we want in a specified order
#
#
def ReportQuestions():
    OrderResp = [""] * 20                   # Be sure we have enough reordering slots
    OrderTally = [0] * 20
    reorder = 0
    x = 0
    TotalTally = 0
    StrongDisagree = 0                      # init the 5 Likert scale response totals
    Disagree = 0
    NoOpinion = 0
    Agree =0
    StrongAgree = 0
    ConfInt = 0                             # not yet known type to do confidence interval calc
    if (len(SurveyFiles) == 1):
        printLine(f"Survey has {TotQuestions} Questions.")
    else:
        printLine(f"Combined the Surveys have {TotQuestions} Total Questions.")
    for q in range(0,TotQuestions):
        ConfInt = 0;                                               # Reset confidence interval for this question
        reorder = 0;                                               # assume answer don't need to be reordered
        for i in range(0,20):
            OrderResp[i] = " "                                     # init report arrays for defined
            OrderTally[i] = 0                                      # with 0 Tallys
        #
        QuestionText[q].replace("\n","\n   ")                      # Indent any newlines 3 spaces
        print (f"Q: {QuestionText[q]}", file=outfh)                # Print Question Text
        #
        #  Add up total responses to this question for later percentage calculations
        #
        TotalTally = QuestionDecline[q]                            # Init Total Tally to # declines
        for i in range(0,QuestionNumResponse[q]):
            TotalTally = TotalTally + QuestionTally[q][i]          # add all answer counts for this question
        #
        #  See if this is a recognized response list
        #
        reorder = 0
        for i in range(0,QuestionNumResponse[q]):
            for x in range(0,5):
                if (QuestionResponse[q][i] == DefAns1[x]):            
                    reorder = 1               # indicate needs reordering against DefAns1 (Likert Scale)
                    break
                if (QuestionResponse[q][i] == DefAns2[x]):
                    reorder = 2               # indicate needs reordering against DefAns2  (Conservative to )
                    break
        if (reorder == 0):
            #
            #  Isn't a response list we know, order it as discovered
            #
            for i in range(0,QuestionNumResponse[q]):
                OrderResp[i] = QuestionResponse[q][i]
                OrderTally[i] = QuestionTally[q][i]
        else:
            #
            #  This is at least partly recognized as Strongly Disagree to Strongly Agree List
            #
            addx = 5                                                    #add 1st unknown here
            if (reorder == 1):
                ConfInt=1
                for x in range(0,5):
                    OrderResp[x] = DefAns1[x]                           # start with defined text order
                for i in range(0,QuestionNumResponse[q]):
                    flag = 0
                    for x in range(0,5):
                        if (QuestionResponse[q][i].lower() == DefAns1[x].lower()):     # this is one of the defined responses
                            OrderTally[x] = QuestionTally[q][i]         #move Tally count to proper slot
                            #
                            #  This is a Likert scale response.  Save the counts for confidence interval calculation
                            #
                            if (DefAns1[x] == "Strongly Disagree"):
                                StrongDisagree = OrderTally[x]
                            if (DefAns1[x] == "Disagree"):
                                Disagree = OrderTally[x]
                            if (DefAns1[x] == "No Opinion"):
                                NoOpinion = OrderTally[x]
                            if (DefAns1[x] == "Agree"):
                                Agree = OrderTally[x]
                            if (DefAns1[x] == "Strongly Agree"):
                                StrongAgree = OrderTally[x]
                            flag = 1
                    if (flag == 0):
                        #
                        #  Not one of the defined answers, put at end
                        #
                        OrderResp[addx] = QuestionResponse[q][i]        # add to next unknown response slot
                        OrderTally[addx] = QuestionTally[q][i]
                        addx += 1                                       # say this slot used
                QuestionNumResponse[q] = addx                           # indicate number of questions to print
            else:
                #
                # This is at least partly recognized as Conservative to Progressive List
                #
                for x in range(0,5):
                    OrderResp[x] = DefAns2[x]                           # start with defined text order
                    OrderTally[x] = 0                                   # and 0 responses
                for i in range(0,QuestionNumResponse[q]):
                    flag = 0
                    for x in range(0,5):
                        QuestionResponse[q][i] = QuestionResponse[q][i].replace("Moderately","Moderate",1)      # Handle change over time
                        if (QuestionResponse[q][i].lower() == DefAns2[x].lower()):     # this is one of the defined responses
                            OrderTally[x] = QuestionTally[q][i]         #move Tally count to proper slot
                            flag = 1
                    if(flag == 0):
                        #
                        #  Not one of the defined answers, put at end
                        #
                        OrderResp[addx] = QuestionResponse[q][i]        # add to next unknown response slot
                        OrderTally[addx] = QuestionTally[q][i]
                        addx += 1                                       # say this slot used
                QuestionNumResponse[q] = addx                           # indicate number of questions to print
        #
        #  Now Report the possibly reordered Array
        #
        Slen = 0                                                # Response Text Character Count
        TCnt = ""                                               # Formatted Tally Count
        RTally = 0
        RText = ""                                              # Formatted Response Text
        TPct = ""                                               # Formatted Percentage text
        BText = ""                                              # Bar Graph Text
        Words = []                                              # Multiword array for multi-line response formatting.
        WdCnt = 0                                               # # of words in multi word array
        Mx = 0                                                  # Multiword array index
        x = 0                                                   # local loop variable
        for i in range(0,QuestionNumResponse[q]):
            Slen = len(OrderResp[i])                            # Get Reponse Text
            RTally = OrderTally[i]
            TCnt = "{0:4d}".format(RTally)                      # Get and Format Response Tally Count
            if (Slen < 26):
                #
                #  Response fits on single line
                #  Build string to print
                #
                RText = OrderResp[i] + (".") * (28-Slen)
                Mx = -1                                         # Indicate not multi-line response
            else:
                #
                #  Response Text is multi-line
                #
                #  Array of words and then build first line of response
                #
                Words = OrderResp[i].split()                    # Split Response Text into Words
                WdCnt = len(Words)                              # and count the words that result
                RText = Words[0]                                # Init Respose String to 1st word
                Slen = len(RText)                               # init String length
                Mx = 1                                          # Point to next word
                while (Slen < 26):
                    if ((Slen + 1 + len(Words[Mx])) < 26):
                        #
                        #  Room to add next word preceeded by a space
                        #
                        RText = RText + " " + Words[Mx]         # add space and next word
                        Slen = len(RText)                       # update string length
                        Mx += 1                                 # point to next word
                    else:
                        #
                        #  Can't add next word, it will be first word of next line
                        #  finish out this line in $Rtext
                        #
                        RText = RText + (".") * (28-Slen)
                        Slen = 30                               # assure exit from while loop
                        #
                        #  Note $Mx is now index of the 1st word in the next line
                        #
            #
            #  Calculate the percentage for this response
            #
            TPct = "{0:4.1f}".format(RTally/TotalTally*100)       # Formatted and rounded % to print
            Slen = int(((RTally/TotalTally*100)/3)+ 0.5)          # get 1/3 of percent (rounded up) = # bar chars
            if (Slen == 0):
                Slen = 1                                        # minimum 1 bar char
            BText = ("■") * Slen                                # Build Graph Bar Text
            print (f"       {RText}{TCnt}   {TPct}%  {BText}", file= outfh)
            #
            #  If multiline response, print rest of the line(s)
            #
            while( (Mx > 0) & (Mx < WdCnt) ):
                RText = Words[Mx]                               # Init Response String to next word
                Slen = len(RText)                               # init String length
                Mx += 1                                         # point to next word
                while ( (Slen < 28) & (Mx < WdCnt) ):
                    if ( ( Slen + 1 + len(Words[Mx])) <= 28 ):
                        #
                        #  Room to add next word preceeded by a space
                        #
                        RText = RText + " " + Words[Mx]         # add space and next word
                        Slen = len(RText)                       # update string length
                        Mx += 1                                 # point to next word
                    else:
                        Slen = 30                               # all that fits on line, exit while loop
                print (f"       {RText}", file = outfh);        # print this line
        #
        #  Add Decline to Answer Line
        #
        RText = "Did Not Answer.............."
        RTally = QuestionDecline[q]
        TCnt = "{0:4d}".format(RTally)
        TPct = "{0:4.1f}".format(RTally/(TotalTally)*100)  # Formatted and rounded % to print
        Slen = int(((RTally/TotalTally*100)/3)+ 0.5)          # get 1/3 of percent (rounded up) = # bar chars
        if (Slen == 0):
            Slen = 1                                        # minimum 1 bar char
        BText = ("■") * Slen                                # Build Graph Bar Text
        print (f"       {RText}{TCnt}   {TPct}%  {BText}", file= outfh)
        if (ConfInt != 0):
            #
            # This is a Likert Scale Response, Do Confidence Interval Calculation
            #
            Disagree = Disagree + StrongDisagree                # Sum Disagree responses
            AgreeNo = NoOpinion + Agree + StrongAgree           # Sum Agree and No Opinion responses
            TotalTally = Disagree + AgreeNo                     # get total responses
            DaFrac = Disagree/TotalTally
            ANFran = AgreeNo/TotalTally
            PlusMinus = math.sqrt(DaFrac*ANFran/TotalTally)
            DALow = DaFrac - (PlusMinus * 1.96)
            DAHi = DaFrac + (PlusMinus * 1.96)
            ALow = 1-DAHi
            AHi = 1-DALow
            #
            #  Calculate 50 character bar graph of confidence interval
            #
            BarLo = int((51*ALow)+0.5)
            BarHi = int((51*AHi)+0.5)
            Leading = BarLo - 1               # of leading dashes
            Trailing = 51 - BarHi             # of Trailing Dashes
            Bar = BarHi - BarLo + 1          # of BAR characters
            CBar = "Disagree "                 # Build confidence bar text
            if (Leading > 0):
                CBar = CBar + "├"                # At least 1 leading char, use end delim
                if (Leading > 1):
                    CBar = CBar + "─" * (Leading-1)   # do rest of leading characters
            CBar = CBar + "■" * Bar
            if (Trailing > 0):
                if (Trailing > 1 ):
                    CBar = CBar + "─" * (Trailing-1)
                CBar = CBar + "┤"
            CBar = CBar + " Agree"
            #substr($CBar,25,1,"┼");
            #
            #  $ALow to $AHi is the confidence interval mean universe response (0-1 scale)
            #
            print (f"\n              95% Confidence Interval: {int(ALow*1000)/10}% - {int(AHi*1000)/10}%", file = outfh)
            print ( CBar , file = outfh)
        print ("\n", file=outfh)
#
#====================================================================================================
#
#--------------------------------------------------------------
# Generate the Day Of Week Report
#--------------------------------------------------------------
def ReportByDayOfWeek():
    global DowAttempts, DowAnswers, DowContacts, DowRefuse, outfh

    sun = ""                       # declare local formatting string variables
    mon = ""
    tue = ""
    wed = ""
    thu = ""
    fri = ""
    sat = ""
    print ("                        Breakdown By Day Of Week", file=outfh)
    print ("                        ------------------------", file=outfh)
    print ("             Sun       Mon       Tue       Wed       Thu       Fri      Sat", file=outfh)
    print ("           -------   -------   -------   -------   -------   -------  -------", file=outfh)
    sun = "{0:4d}".format(DowAttempts[0])
    mon = "{0:4d}".format(DowAttempts[1])
    tue = "{0:4d}".format(DowAttempts[2])
    wed = "{0:4d}".format(DowAttempts[3])
    thu = "{0:4d}".format(DowAttempts[4])
    fri = "{0:4d}".format(DowAttempts[5])
    sat = "{0:4d}".format(DowAttempts[6])
    print (f" Attempts   {sun}      {mon}      {tue}      {wed}      {thu}      {fri}     {sat}",file=outfh)
    sun = "{0:4d}".format(DowAnswers[0])
    mon = "{0:4d}".format(DowAnswers[1])
    tue = "{0:4d}".format(DowAnswers[2])
    wed = "{0:4d}".format(DowAnswers[3])
    thu = "{0:4d}".format(DowAnswers[4])
    fri = "{0:4d}".format(DowAnswers[5])
    sat = "{0:4d}".format(DowAnswers[6])
    print (f"  Answers   {sun}      {mon}      {tue}      {wed}      {thu}      {fri}     {sat}", file=outfh)
    sun = "{0:4d}".format(DowContacts[0])
    mon = "{0:4d}".format(DowContacts[1])
    tue = "{0:4d}".format(DowContacts[2])
    wed = "{0:4d}".format(DowContacts[3])
    thu = "{0:4d}".format(DowContacts[4])
    fri = "{0:4d}".format(DowContacts[5])
    sat = "{0:4d}".format(DowContacts[6])
    print (f" Contacts   {sun}      {mon}      {tue}      {wed}      {thu}      {fri}     {sat}", file=outfh)
    sun = "{0:4d}".format(DowRefuse[0])
    mon = "{0:4d}".format(DowRefuse[1])
    tue = "{0:4d}".format(DowRefuse[2])
    wed = "{0:4d}".format(DowRefuse[3])
    thu = "{0:4d}".format(DowRefuse[4])
    fri = "{0:4d}".format(DowRefuse[5])
    sat = "{0:4d}".format(DowRefuse[6])
    print (f"   Refuse   {sun}      {mon}      {tue}      {wed}      {thu}      {fri}     {sat}", file=outfh)
    for x in range(0,7):
        if (DowAttempts[x] == 0):
            DowAttempts[x] = 0.1;             # Avoid divide by zero
    sun = "{0:4.1f}".format(DowAnswers[0]/DowAttempts[0]*100)        # Format Answer %
    mon = "{0:4.1f}".format(DowAnswers[1]/DowAttempts[1]*100)
    tue = "{0:4.1f}".format(DowAnswers[2]/DowAttempts[2]*100)
    wed = "{0:4.1f}".format(DowAnswers[3]/DowAttempts[3]*100)
    thu = "{0:4.1f}".format(DowAnswers[4]/DowAttempts[4]*100)
    fri = "{0:4.1f}".format(DowAnswers[5]/DowAttempts[5]*100)
    sat = "{0:4.1f}".format(DowAnswers[6]/DowAttempts[6]*100)
    print (f"  Answer%   {sun}%     {mon}%     {tue}%     {wed}%     {thu}%     {fri}%    {sat}%", file=outfh)
    sun = "{0:4.1f}".format(DowContacts[0]/DowAttempts[0]*100)        # Format Complete %
    mon = "{0:4.1f}".format(DowContacts[1]/DowAttempts[1]*100)
    tue = "{0:4.1f}".format(DowContacts[2]/DowAttempts[2]*100)
    wed = "{0:4.1f}".format(DowContacts[3]/DowAttempts[3]*100)
    thu = "{0:4.1f}".format(DowContacts[4]/DowAttempts[4]*100)
    fri = "{0:4.1f}".format(DowContacts[5]/DowAttempts[5]*100)
    sat = "{0:4.1f}".format(DowContacts[6]/DowAttempts[6]*100)
    print (f"Complete%   {sun}%     {mon}%     {tue}%     {wed}%     {thu}%     {fri}%    {sat}%", file=outfh)
    for x in range(0,7):
        if (DowAnswers[x] == 0):
            DowAnswers[x] = 0.1             # Avoid divide by zero
    sun = "{0:4.1f}".format(DowRefuse[0]/DowAnswers[0]*100)        # Format Answer %
    mon = "{0:4.1f}".format(DowRefuse[1]/DowAnswers[1]*100)
    tue = "{0:4.1f}".format(DowRefuse[2]/DowAnswers[2]*100)
    wed = "{0:4.1f}".format(DowRefuse[3]/DowAnswers[3]*100)
    thu = "{0:4.1f}".format(DowRefuse[4]/DowAnswers[4]*100)
    fri = "{0:4.1f}".format(DowRefuse[5]/DowAnswers[5]*100)
    sat = "{0:4.1f}".format(DowRefuse[6]/DowAnswers[6]*100)
    print (f"  Refuse%   {sun}%     {mon}%     {tue}%     {wed}%     {thu}%     {fri}%    {sat}%", file=outfh)
    print ("\n", file=outfh)
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
    global datadir, rptfile, svyfile, svydir, qfile, sframe, printFileh, outfh
    global Headings, NumSvyRows, NumSvyCols, SurveyTitle, DateCol
    global VolFirstCol, VolLastCol, PhoneCol, ResponseCol
    global SurveyFiles, ResponseCode, ProgName
    global NumVolunteer, VolunteerName, VolunteerAttempts, VolunteerAnswers
    global VolunteerContacts, VolunteerRefuse, VolunteerPartial
    global NumAttempts, NumAnswers, NumContacts, NumRefuse, NumPartial
    global NumPhones, PhoneNumber, PhoneDuplicates
    global NumResponse, ResponseCode, ResponseCnt
    global DowIndex, DowText, DowAttempts, DowAnswers, DowContacts, DowRefuse
    global NumQuestions, QuestionText, QuestionNumResponse, QuestionResponse
    global QuestionTally, QuestionDecline, TotQuestions, QuestionIndex
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
        print ("I/O error({0}): {1}".format(e.errno, e.strerror))
        exit(2)
    except: #handle other exceptions such as attribute errors
        print ("Unexpected error:", sys.exc_info()[0])
        exit(2)
    #
    ProgName = sys.argv[0][0:-3].upper()            # Stash program Name (minus .py) in upper case for PrintLine
    args(sys.argv[1:])                              # Get command line arguments if any
    #
    #--------------------------------------------------------------------
    #  Find and Select the Survey Files this program will report
    #--------------------------------------------------------------------
    #
    if (svyfile != ""):
        SurveyFiles.append(svyfile)                                     # infile overrides any selection of files
    else:
        dir = os.getcwd()                                               # Get our current working directory
        if (svydir != ""):
            dir = f"{dir}/{svydir}"                                     # if a Survey Subdirectory specified add it to path
            dir = os.path.normpath(dir)                                 # normalize to Windows or Unix path syntax
        printLine(f"Survey Directory = {dir}")
        if (os.path.isdir(dir) == False):
            printLine (">>> Directory does not exist --- aborting!\n")
            exit(2)
        allfiles = [f for f in listdir(dir) if isfile(join(dir, f))]    # get list of files in survey directory
        sfiles = []
        #
        #  Build list of only survey file names
        #
        for x in range(0,len(allfiles)):                                # loop through file names
            temp = allfiles[x]                                          # get next file name
            temp = temp.lower()                                         # force name lower case
            if (temp.find(".xlsx") == -1):
                continue
            temp = temp.replace(".xlsx","")                             # it's a .xlsx file, remove file extension
            flag = 0                                                    # assume good survey file
            if (temp[0 : 5] == "i-20p" ):
                flag = 1
            if ((temp[0] == "p") and (temp[2] == "-")):                 # if name starts with i-20p or pn- it's a survey
                flag = 1
            if(flag == 0):
                if(temp[0:14] == "surveyresponse"):
                    SurveyFiles.append(f"{dir}/{allfiles[x]}")          # This file is selected add to survey file name list
                continue                                                # ignore non-selected file
            else:  
                if(qualify_file(temp) == 0):                            # screen any command line selection criteria
                    SurveyFiles.append(f"{dir}/{allfiles[x]}")          # This file is selected add to survey file name list
    #
    #####################################################################################
    #
    #  The array SurveyFiles is now a list of zero or more survey files to report
    #
    ######################################################################################
    #    
    if (len(SurveyFiles) == 0):
        printLine ("No input files found -- nothing to report...")
        exit (2)
    #
    #  Now process all selected survey files in list SurveyFiles
    #
    for svyfile in SurveyFiles:                                 # get next file name to process
        if (svyfile == SurveyFiles[0]):
            #
            #  check file name to see if .csv or .xlsx and make each
            #  read use either read_csv or read_excel as needed to allow full
            #  flexibility in input files.
            #
            printLine (f"Loading {svyfile} ... ")
            if (svyfile[-4:] == ".csv"):
                sframe = pd.read_csv (svyfile,low_memory=False, encoding='latin-1')     #  Read .csv survey file into dataframe "sframe"
            else:
                sframe = pd.read_excel (svyfile)                    #  Read .xls or .xlsx survey file into dataframe "sframe"
            ec = init_spreadsheet()                                 #  initialize spreadsheet required variables
            if(ec != 0):
                printLine (f"Spreadsheet not in expected survey format... Error Code {ec} skipping")
                continue
            printLine (f"... Survey data loaded - {NumSvyRows} Rows, {NumSvyCols} columns")
            #
            if ("SurveyResponse" in svyfile):
                if (qualify_file(SurveyTitle) != 0):                # use title from spreadsheet data to see if it is selected
                    printLine (f"... Survey name {SurveyTitle} not selected ... skipping")
                    continue
            #
            #-----------------------------------------------------------------------------------
            #  Initialize Question Discovery and Tally Variables for FIRST spreadsheet in list
            #-----------------------------------------------------------------------------------
            #
            NumQuestions = NumSvyCols - FirstQCol - ColsAfterQ  # Calculate The Number of Questions in this spreadsheet
            TotQuestions = NumQuestions                         # first spreadsheet, this is also all the questions so far
            QuestionIndex =[]                                   # reset question index array
            for q in range(0,NumQuestions):
                QuestionIndex.append(q)                         # for first file, index = 1, 2, ....
                QuestionText.append(Headings[q+FirstQCol])      # Build Question Text Table
                QuestionResponse.append([])                     # create correct number of pointer entries
                QuestionTally.append([])                        # for reference pointer arrays
                QuestionNumResponse.append(0)                   # No Responses so far
                QuestionDecline.append(0)                       # no decline to answers so far
            printLine(f"... Initialized for {NumQuestions} Questions ...")
        else:
            #----------------------------------------------------------------------
            #  Open and Load 2nd to nth survey file for -survey option
            #----------------------------------------------------------------------
            #
            # Open & Load next -survey spreadsheet into array pointed to by $bookdata
            #
            new = 0                                                 # assume no new questions
            printLine (f"Loading {svyfile} ... ")
            if (svyfile[-4:] == ".csv"):
                sframe = pd.read_csv (svyfile,low_memory=False)     #  Read .csv survey file into dataframe "sframe"
            else:
                sframe = pd.read_excel (svyfile)                    #  Read .xls or .xlsx survey file into dataframe "sframe"
            ec = init_spreadsheet()                                 #  initialize spreadsheet required variables
            if(ec != 0):
                printLine (f"Spreadsheet not in expected survey format... Error Code {ec} skipping")
                continue
            printLine (f"... Survey data loaded - {NumSvyRows} Rows, {NumSvyCols} columns")
            #
            if ("SurveyResponse" in svyfile):
                if (qualify_file(SurveyTitle) != 0):                # use title from spreadsheet data to see if it is selected
                    printLine (f"... Survey name {SurveyTitle} not selected ... skipping")
                    continue                                        # get basic stuff and verify this is survey spreadsheet
            #
            #  Initialize Question Discovery and Tally Variables for 1st spreadsheet
            #
            NumQuestions = NumSvyCols - FirstQCol - ColsAfterQ      # Calculate The Number of Questions in this spreadsheet
            QuestionIndex = []                                      # reset question index array
            for i in range(0,NumQuestions):
                QuestionIndex.append(-1)                            # init QuestionIndex array with -1 in each slot
            for i in range(0,NumQuestions):                         # process these questions to see if new or not
                temp = -1                                           # init as new question
                for q in range(0,TotQuestions):
                    if (Headings[i+FirstQCol] == QuestionText[q]):
                        QuestionIndex[i] = q                        # repeat question, point to it in parallel arrays
                        temp=0                                      # flag repeat question
                if (temp == -1):
                    new += 1                                        # count new questions
                    QuestionIndex[i] = TotQuestions                 # index for this question
                    TotQuestions += 1                               # one more total question
                    QuestionText.append(Headings[i+FirstQCol])      # Build Question Text Table
                    QuestionResponse.append([])                     # Add Response empty list
                    QuestionTally.append([])                        # Add Tally empty list
                    QuestionNumResponse.append(0)                   # No Responses so far
                    QuestionDecline.append(0)                       # No decline to answers so far
            for i in range(0,NumQuestions):
                if (QuestionIndex[i] == -1):                        # safety check
                    printLine ("Internal Failure in 2nd to nth Survey File load ... aborting")
                    exit (2)
            if (new > 0):
                printLine(f"... Initialized for {NumQuestions} Questions ({new} new) ...")
            else:
                printLine(f"... Initialized for {NumQuestions} Questions ...")
        #
        #-----------------------------------------------------------
        #  Next survey file loaded initialize date range of survey
        #-----------------------------------------------------------
        #
        #  Fix up anomalies in data file we've encountered along the way
        #
        sframe = sframe.replace(np.nan, '', regex=True)         #  make any nans into '' anywhere in the data frame
        if(sframe.columns[0] != "Title"):                       # Fix corrupted 1st Column name we see in some i360 files
            CorruptHdr = sframe.columns[0]
            sframe.rename(columns={CorruptHdr: 'Title'}, inplace=True)
        #
        rdate = sframe.iloc[0][DateText]                        # get response date from first row
        if (isinstance(rdate,str)):
            if(rdate[4] == "-"):
                #
                #  Date string is yyyy-mm-dd format
                #
                rdate = datetime.strptime(rdate,"%Y-%m-%d")     # Text date yyyy-mm-dd, convert to datetime object
            else:
                #
                #  date string is of one of the following formats:
                #
                #  mm/dd/yyyy,  m/dd/yyyy, mm/d/yyyy, m/d/yyyy
                #  mm/dd/yy,  m/dd/yy, mm/d/yy, m/d/yy
                # 
                # Make sure it's mm/dd/yyyy if it's not already 
                #  
                if (rdate[1] == '/'):
                    rdate = '0' + rdate                         # single digit month, make 0n
                if (rdate[4] == '/'):
                    rdate = rdate[0:3] + '0' + rdate[3:]        # single digit day, make 0n
                if (len(rdate) == 8):
                    rdate = rdate[0:6] + '20' +rdate[6:]        # two digit year make 20nn
                rdate = datetime.strptime(rdate,"%m/%d/%Y")     # Text date mm/dd/yyyy, convert to datetime object
        else:
            rdate = rdate.to_pydatetime()                       # convert pandas timestamp to datetime object
        svystart = rdate                                        # init start and end dates as first row's date
        svyend = rdate
        #
        #  Now analyze data for row 1-n
        #
        printLine ("... Analyzing Survey data ...")
        for row in range(0,NumSvyRows):
            if ((row % 100) == 0):
                print(f"Processing Row {row}\r", end="")
            rlist = list(sframe.iloc[row])                      # get row data as a list
            rdate = rlist[DateCol]                              # get response date from next row
            #
            #  If there are any -pers persuasion qualifier options, see if this row qualifies
            #
            if (NumPersOpts > 0):
                flag = 0                                        # Assume row should be skipped
                for q in range(0,NumQuestions):
                    for ii in range(0,NumPersOpts):
                        if (rlist[FirstQCol+q] == PersOpts[ii]):
                            flag = 1                            # this row qualifies, process it
                            break                               # end lookup loop
                if (flag == 0):
                    continue                                    # skip processing this row if it doesn't qualify
            #
            NumAttempts += 1                                    # Every Row is an attempt
            #
            # Process Date in either format (depending on whether .csv or .xlsx) file
            #
            if (isinstance(rdate,str)):
                if(rdate[4] == "-"):
                    rdate = datetime.strptime(rdate,"%Y-%m-%d") # Text date yyyy-mm-dd, convert to datetime object
                else:  
                    if (rdate[1] == '/'):
                        rdate = '0' + rdate                         # single digit month, make 0n
                    if (rdate[4] == '/'):
                        rdate = rdate[0:3] + '0' + rdate[3:]        # single digit day, make 0n
                    if (len(rdate) == 8):
                        rdate = rdate[0:6] + '20' +rdate[6:]        # two digit year make 20nn
                    rdate = datetime.strptime(rdate,"%m/%d/%Y") # Text date mm-dd-yyyy, convert to datetime object
            else:
                rdate = rdate.to_pydatetime()                   # date is pandas Timestamp, convert to datetime object
            if (rdate < svystart):
                svystart = rdate                                # older than start, so new start date
            if (rdate > svyend):
                svyend = rdate                                  # newer than end date, so new end date
            DowIndex = rdate.weekday()                          # get day of week from datetime object (0 = Monday, 1 = tuesday etc)
            if (DowIndex == 6):                                 # Convert index to 0 = Sunday, 1 = Monday, etc.)
                DowIndex = 0                                    # Sunday day of week index = 0
            else:
                DowIndex += 1                                   # Mon - Sat index (1=Monday, 2 = Tuesday etc.,)
            DowAttempts[DowIndex] += 1                          # log attempts by day of week
            #
            #  See if Unique Phone Number, if so add to table.  If Not, count duplicates
            #
            if(rlist[PhoneCol] != ""):                          # don't check blank phone numbers
                if rlist[PhoneCol] in PhoneNumber:
                    PhoneDuplicates += 1                        # this is a duplicate Phone Number
                else:
                    PhoneNumber.append(rlist[PhoneCol])         # Unique Phone Number so far, add to table
                    NumPhones += 1                              # Count Unique Numbers Called
            #
            # Build ResponseText array and log count for each text response in ResponseCnt array
            #
            rtext = rlist[ResponseCol]                          # fetch response text
            flag=0                                              # assume new response
            for y in range(0,NumResponse):
                if (rtext == ResponseCode[y]):
                    ResponseCnt[y] += 1
                    flag = 1                                    # repeat response, don't add to tables
                    break
            if ( flag == 0):
                NumResponse += 1                                # new response, say one more found
                ResponseCode.append(rtext)                      # add this response to list
                ResponseCnt.append(1)                           # with 1 hit so far
            #
            # Build Volunteer Name Array and Counts by Volunteer
            #
            try:
                Vname = rlist[VolFirstCol] + " " + rlist[VolLastCol] # get volunteer name
            except:
                print(f"Error on Row {row}")
                print(rlist)
            try:
                Vx = VolunteerName.index(Vname)                 # find this Volunteer in Table
                VolunteerAttempts[Vx] += 1                      # Bump Attempts
            except ValueError:
                VolunteerName.append(Vname)                     # New Name, add to table
                VolunteerAttempts.append(1)                     # init parallel Volunteer lists
                VolunteerAnswers.append(0)
                VolunteerContacts.append(0)
                VolunteerPartial.append(0)
                VolunteerRefuse.append(0)
                Vx = NumVolunteer
                NumVolunteer += 1                               # count new entry
            #
            #  Vx now points to the volunteer for this record
            #
            rtext = rlist[ResponseCol]                               # fetch response text
            if (rtext == Take):
                VolunteerAnswers[Vx] += 1                   # Survey = Answer & Taken
                VolunteerContacts[Vx] += 1
                NumAnswers += 1                             # Add to total Answers
                NumContacts += 1                            # add to total Surveys Taken
                DowContacts[DowIndex] += 1                  # also by day of week
                DowAnswers[DowIndex] += 1                   # Answers too
                if (doquestions(rlist) != 0):               # process question answers
                    NumPartial += 1                         # only some questions answered
                    VolunteerPartial[Vx] += 1               # say this was a partial survey
            if (rtext == Refuse):
                VolunteerAnswers[Vx] += 1
                VolunteerRefuse[Vx] += 1
                NumAnswers += 1                             # Add to total Answers
                NumRefuse += 1                              # and total Refuses
                DowRefuse[DowIndex] += 1                    # also by day of week
                DowAnswers[DowIndex] += 1                   # Answers too
            #
            #   ----- End Spreadsheet row processing loop ------
            #
        #
        #  All survey file selected are now processed
        #
        # data is now processed into the tables
        #
        pstart = svystart.strftime("%m/%d/%y")              # format start adn stop dates for printing
        pend = svyend.strftime("%m/%d/%y")
        printLine (f"... Survey ran from {pstart} to {pend}")
    #
    #
    #
    #  >>>>>>>>>>>>> All Survey Files are loaded, do the reporting <<<<<<<<<<<<<<<<<<<<
    #
    #  See if any question text should be substituted from an optional question text file
    #
    numsub = 0
    if (qfile != ""):
        #
        #  question text substitution file specified, process it
        #
        printLine(f"Substituting question text from {qfile}")
        if (qfile[1] != ":"):
            dir = os.getcwd()                                               # Get our current working directory
            qfile = os.path.join(dir, qfile)                                # fully qualify relative file path
        try:
            qfileh = open(qfile,"r")
            for row in qfileh:
                #
                #  find matching question and substitute text
                #  consider question a match if first 20 charactrers match
                #
                i=len(row)
                if (i < 3):
                    continue                    # skip blank lines
                if (i > 50):
                    i = 50                      # check max of 50 characters for match
                for q in range(0,TotQuestions):
                    if ( row[0:i].lower() == QuestionText[q][0:i].lower() ):
                        QuestionText[q] = row   # substitute this text
                        numsub += 1
                        break
            qfileh.close
        except IOError as e:
            printLine (f"Could not open question file '{qfile}' reason: {e.strerror}")
            printLine ("... Continuing with survey file question text...")
        #
        if (numsub > 0):
            printLine (f"... {numsub} questions had text substituted...")
    #
    #  Create the output report file
    #  If created from a SurveyRespone file directly, then output the report to a file named from 
    #    the Report Title and placed in the directory specified by 'datadir'/reports 
    #    also, rename the SurveyResponse file to a file named from Report Title and move it
    #      the directory specified by 'datadir'/surveys
    #  else create the report file in the current working directory with the name report.txt
    #
    title = sframe.iloc[0]["Title"]
    if(rptfile[2] != ":"):  Dir = os.getcwd()                                       # Output to our current working directory
    rptfile = os.path.join(Dir,rptfile)                     # form full temp file name
 
    if ("SurveyResponse" in rptfile):
        yyyymmdd = datetime.today().date().isoformat().replace("-", "")      # set automatic report file name
        rptfile = datadir + "reports/" + title + "-" + yyyymmdd + ".txt"      # set default report file name
        # now move the survey file to the 'datadir'/surveys directory with a new name
        svyfile2 = datadir + "surveys/" + title + "-" + yyyymmdd + ".xlsx"    # rename and move survey file
        shutil.move(svyfile, svyfile2)
        
    printLine (f"... Generating Reports to file {rptfile}")
    try:
        outfh = open(rptfile,"w",encoding='utf-8')
    except IOError as e:
        print ("Cannot open report file, error({0}): {1}".format(e.errno, e.strerror))
        exit(2)
    except: #handle other exceptions such as attribute errors
        print ("Cannot open report file, Unexpected error:", sys.exc_info()[0])
        exit(2)
    #
    #  Write the report to the file
    #
    ReportAttempts(pstart,pend)
    ReportCallStatus()
    ReportByVolunteer()
    ReportByDayOfWeek()
    ReportQuestions()
    #
    # Close the output report file
    #
    outfh.close()
    #
    #  End of program
    #
    EndTime = time.time()
    printLine (f"Total Elapsed time is {int((EndTime - StartTime)*10)/10} seconds.\n")
    printFileh.close()
    return (0)

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
