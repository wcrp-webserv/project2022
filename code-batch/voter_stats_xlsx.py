
#-----------------------------------------------------------------------------------
#                               voter_screen.py                                     
#                                                                                   
#  purpose: currently creates some statistics by analyzing a base.csv type file and
#   input:  analyze a file in the form of base.csv from nvvoter2.pl
#           filters: "type": ad or sd 
#           "district" number of ad or sd
#           "ALL"
#   output: various counts and averages from that file
#                
#----------------------------------------------------------------------------------- 
import numpy as np
import argparse
import errno
import os
import xlsxwriter
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
voted = ['BR', 'EV', 'MB', 'PP', 'PV']


#
# get command line arguments
# voter-screen.py" sd5.csv SD 5 
parser = argparse.ArgumentParser(description='Process the inputs')
parser.add_argument("-infile", type=str, help="the name of the file to process")
parser.add_argument("-dist", type=str, help="ad or sd")
parser.add_argument("-id", type=str, help="number of district")


args = parser.parse_args()
# echo the command line
print (args.infile, args.dist, args.id)

#
# function: to divide x/y and return result, protects from divide by 0
#
def div0(x,y):
    try:
        return x/y
    except ZeroDivisionError:
        return 0

#
# function: process a row of the base.csv tabl
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
    
    if (row['Party'] == 'Democrat'):
        democrats += 1
        if vote in voted:
            anyvote  +=  1
            demvote  +=  1
    elif (row['Party'] == 'Republican'):
        republicans  +=  1
        if vote in voted:
            anyvote  +=  1
            repvote  +=  1
    elif (row['Party'] == 'Non-Partisan'):
        nonps  +=  1
        if vote in voted:
            anyvote  +=  1
            nonpvote = nonpvote+1
    elif (row['Party'] == 'Independent American Party'):
        iaps  +=  1
        if vote in voted:
            anyvote  +=  1
            iapvote  +=  1
    else:
        other  +=  1
        if vote in voted:
            anyvote  +=  1
            othvote   +=  1
        return 0

def main():
    column = ""
    if args.dist == "SD":
        column = "SenDist"
    elif args.dist == "AD":
        column = "AssmDist"
    analyze_file = args.infile
    mycwd = os.getcwd()
    rownum = 0
    allvoters = 0


    # read data source 
    with open(analyze_file, mode='r', encoding='ISO-8859-1') as csv_file:
        csv_reader = DictReader(csv_file)
        print ("opened file: {}".format(analyze_file))
        #voted = ['BR', 'EV', 'MB', 'PP', 'PV']

        # create document
        try:
            for row in csv_reader:
                if rownum == 0:
                    rownum = 1
                    continue
                rownum  += 1
                if (args.dist == 'ALL'):
                    pass
                elif (row[column] != args.district):
                    continue
                allvoters += 1
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

    #exit(0)
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook("stats.")
    worksheet = workbook.add_worksheet()

    # set workbook print properties
    worksheet.set_landscape()                                           # set to print in landscape orientation
    worksheet.set_paper(5)                                              # 5 for legal paper  (1 for Letter Paper)
    worksheet.set_margins(left = 0.7, right = 0.7, top = 0.75, bottom = 0.75)   # set print margins to Excel "normal"
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
    x = 0
    outheader = ["CountyID",
            "First",
            "Last",
            "Middle",
            "Score"
]
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
    row = 0
    for x in range(count):
        outrow = list(extract.iloc[x])                              # get next row of dataframe as list
        i = 0
        for item in outrow:
            if (item != item):                                      # Make any NAN cells null strings
                outrow[i] = ""
            i += 1
        party = outrow[15]                                          # Fetch Party of this voter
        if (party == "Democrat"):
            numDem=numDem + 1                                       # Add to number of Democrat Voters
            continue                                                # don't write out record
        row = row+1                                                   # write to next row
        worksheet.write_number (row, 0, outrow[0], fmt_right)       # CountyID
        worksheet.write_string (row, 1, outrow[7], fmt_left)        # First
        L = len(outrow[7])
        if (L > MaxFirst):
            MaxFirst = L                                            # longest so far
        worksheet.write_string (row, 2, outrow[8], fmt_left)        # Last
        L = len(outrow[8])
        if (L > MaxLast):
            MaxLast = L                                             # longest so far
        worksheet.write_string (row, 3, outrow[9], fmt_left)        # Middle
        L = len(outrow[9])
        if (L > MaxMiddle):
            MaxMiddle = L                                           # longest so far
        phone = outrow[11]
        if (phone.isdigit()):
            worksheet.write_number (row, 4, int(phone), fmt_right)  # Phone is numeric
        elif (phone == ""):
            worksheet.write_blank(row, 4, None)                     # No Phone Number
        else:
            worksheet.write (row, 4, phone, fmt_right)              # Phone not numeric
        worksheet.write (row, 5, outrow[14], date_format)           # Regdate
        worksheet.write_string (row, 6, party, fmt_left)            # Party
        if (party == "Republican"):
            numRep = numRep + 1                                       # Add to # Republican Voters
        else:
            numOth = numOth + 1                                     # Add to number of "Other Party" Voters
        snum = outrow[16]
        if (snum == ""):
            worksheet.write_blank(row, 7, None)                     # Street Number is blank
        else:
            worksheet.write_number (row, 7, int(snum), fmt_right)   # Street Number is numeric
        SName = outrow[17]
        if (SName == ""):
            worksheet.write_blank(row, 8, None)                     # Street Name Blank
        else:
            worksheet.write (row, 8, outrow[17], fmt_left)          # Street Name
        L = len(outrow[17])
        if (L > MaxStreet):
            MaxStreet = L                                           # longest so far
        Days = outrow[24]
        if (Days == ""):
            print("RegDays Blank in Row " + str(row+1) + " Precinct " + str(PctNum))
            worksheet.write_blank(row, 9, None)
        else:
            worksheet.write_number (row, 9, outrow[24], fmt_right)     # Reg Days
        Age = outrow[25]
        if (Age == ""):
            print("Age Blank in Row " + str(row+1) + " Precinct " + str(PctNum))
            worksheet.write_blank(row, 10, None)
        else:
            worksheet.write_number (row, 10, Age, fmt_right)        # Age
        worksheet.write_string (row, 11, outrow[53], fmt_left)      # Likely To Vote
        LastVote = "Never"
        for i in range(20):
            if (outrow[26+i] != ""):
                LastVote = basehead[26+i]
                break
        LastVote = LastVote[0:6] + "20" + LastVote[6:8]             # truncate to date only and expand year to 4 digits
        worksheet.write_string (row, 12, LastVote, date_format)     # latest Election Voted In
        worksheet.write_number (row, 13, outrow[54], fmt_right)     # Score
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
    print (f"Total Elapsed time is {int((EndTime - StartTime)*10)/10} seconds.\n")
    exit(0)


# --------------------------------------------------
if __name__ == '__main__':
    main()

    



    # count by party 
    
        