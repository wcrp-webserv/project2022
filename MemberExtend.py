#************************************************************************************
#                               MemberExtend.py                                     *
#                                                                                   *
#  Input is CC Member Excel spreadsheet, NVVoter.pl Processed Secretary of          *
#  State base.csv file and precinct-to-district cross reference .csv file.          *
#  Also input is a NickName to Given Name Spreadsheet to aid in name matching.      *
#                                                                                   *
#  Output is an expanded CC member csv file that adds the districts each member     *
#  votes for, their voting propensity and history, along with age and               *
#  date registered to vote to the original member file information.                 *
# *********************************************************************************** 

import pandas as pd
import numpy as np
import sys, getopt


Sosfile = "base.csv"                       # Secretary of State Data with voting results combined
Pctfile = "adall-precincts-20jun.csv"      # Precinct to district file
Memfile = "Central Committee 6-9-2020.xls" # Central Committee Member file
outfile = "extract.csv"                    # output extended member file
nickfile = "NickNameList.xls"              # Nickname to Fullname file
base=""
mbrow=[]

#*******************************************************
#                                                      *
#  Routine to get command line arguments (if any)      *
#                                                      *
#*******************************************************
def args(argv):
   global Sosfile
   global Pctfile
   global Memfile
   global outfile
   try:
      opts, args = getopt.getopt(argv,"h:o:s:m:p:",["sosfile=","outfile=", "memfile=", "pctfile="])
   except getopt.GetoptError:
      print('test.py -s <Sosfile> -p <Pctfile> -m <Memfile> -o <outfile>')
      sys.exit(2)
   for opt, arg in opts:
      if opt == '-h':
         print('test.py -s <Sosfile> -p <Pctfile> -m <Memfile> -o <outfile>')
         sys.exit()
      elif opt in ("-s", "--sosfile"):
         Sosfile = arg
      elif opt in ("-m", "--memfile"):
         Memfile = arg
      elif opt in ("-p", "--pctfile"):
         Pctfile = arg
      elif opt in ("-o", "--outfile"):
         outfile = arg
   print("Input files:")
   temp = '   SOS data file is "' + Sosfile + '"'
   print(temp)
   temp = '   Pct data file is "' + Pctfile + '"'
   print(temp)
   temp = '   Mem data file is "' + Memfile + '"'
   print(temp)
   temp = 'Output file is "' + outfile + '"'
   print(temp)
   print("")

#*****************************************************
#  Build .csv string from a list of column values    *
#                                                    *
#      input is list of column values                *
#      returns .csv string for this list             *
#*****************************************************
def buildcsv(row):
   chars = set("""'",""")
   prow=str(row[0])                             # get 1st column
   if any((c in chars) for c in prow):
         prow = '"'+prow+'"'                    # quote it if needed'
   for x in range(1, len(row)):
      cell=str(row[x])                          # get next column
      if cell == 'nan':
         cell = ""                              # handle empty column
      if any((c in chars) for c in cell):
         cell = '"'+cell+'"'                    # quote if needed
      prow=prow + ',' + str(cell)               # add to .csv string
   prow=prow + '\n'                             # terminate .csv string with newline
   return(prow)

#****************************************************
#                                                   *
#   Lookup Name in SOS base.csv file and return     *
#   matching SOS record.  If no match, return       *
#   SOS record with empty fields in each column     *
#                                                   *
#****************************************************
def lookupbase(First, Last, Middle, Given):
   global base
   global mbrow
   mbx=""
   try:
      if (Middle == ""):
         mbx = base.loc[(base['Last'].str.lower() == Last) & (base['First'].str.lower() == First)]
         if((mbx.empty) and (Given != "")):
            # Nickname failed, try given name form cross reference file
            mbx = base.loc[(base['Last'].str.lower() == Last) & (base['First'].str.lower() == Given)]
         if(mbx.empty):
            # Still Failed, try middle name as First Name
            mbx = base.loc[(base['Last'].str.lower() == Last) & (base['Middle'].str.lower() == First)]
      else:
         Search with middle initial in addition to first and last name
         mbx = base.loc[(base['Last'].str.lower() == Last) & (base['First'].str.lower() == First) & (base['Middle'].astype(str).str[0].str.lower() == Middle)]

      mbrow = mbx.values.flatten().tolist()           # fetch row into list
      if (len(mbrow) < 50):
         print("Error finding base row for " + First + " " + Middle + " " + Last)
         mbrow = [""] * 55
   except KeyError:
      print("Key Error finding base record for " +mrow[2] + " " + mrow[1])


#**********************************************
#    M A I N   P R O G R A M   S T A R T      *
#**********************************************
#
def main():
   global base
   args(sys.argv[1:])                                 #  Get command line arguments if any
   #
   #  Can expand this to check file name to see if .csv or .xls and make each
   #  read use either read_csv or read_excel as needed to allow full
   #  flexibility in input files.
   #
   #  For now, output file is always a .csv file.
   #
   base = pd.read_csv (Sosfile,low_memory=False)      #  Read SOS base.csv file into DataFrame "base"
   member = pd.read_excel(Memfile, sheet_name=0)      #  Read Central Committee excel file (1st sheet) into Dataframe "member"
   pct = pd.read_csv (Pctfile,low_memory=False)       #  Read in the Precinct to district conversion file into DataFrame "pct"
   nick = pd.read_excel(nickfile, sheet_name=0)       # Read Nickname ot Full Name file into DataFrame "nick"
   #
   # get lists of columnn lbels from teh three input files
   #
   basehead=list(base.columns)                        # get SOS data column labels
   cchead=list(member.columns)                        # get Member file column labels
   pcthead=list(pct.columns)                          # get precinct to district column labels
   #
   #  Create Header and write as 1st line to .csv output file
   #
   outhead = cchead.copy()                            # start with member file header
   outhead.extend(pcthead[1:6])                       # add Senate, Assembly, BOE, Regent and Commisioner districts
   outhead.append(basehead[14])                       # add date registered to vote
   outhead.append(basehead[25])                       # add Age
   outhead.extend(basehead[46:55])                    # add voting propensity caclulations
   outhead.extend(basehead[26:46])                    # add actual voting method history
   #
   try:
      out = open(outfile,'w',encoding = 'utf-8')      # open output file and write header line to it.
      out.write(buildcsv(outhead))
   except:
      print("Error opening output file...aborting\n")
      exit(0)
   #
   #  Initialization is done.  Now:
   #    1. read in each member from member file
   #    2. locate matching data in SOS data
   #    3. locate matching record in Precinct-to-district file
   #    4. build expanded member record with added fields as per header above
   #    5. write to output .csv file
   #
   print("... All Files Opened and Loaded...\n")

   for x in range(0 , len(member['Last Name'])):
      mrow = list(member.iloc[x])                        # fetch Member in row into list
      #
      #  Following is an attempt to handle names enetered in funky way in member file
      #
      First = mrow[2].strip().lower()                    # get first name with leading/trailing spaces removed lower cased
      First = First.replace('.', '')                     # remove any periods
      Last = mrow[1].strip().lower()                     # get last name
      Last = Last.replace('.', '')                       # remove any periods
      Middle = ""                                        # assume no middle name
      spacex = First.find(" ")                           # get index of any space in First Name
      if (spacex != -1):
         if (len(First) == 3):
            First = First.replace(' ','')
         else:
            Middle = First[spacex+1]                        # get possible middle Character
            if ((Middle == '"') or (Middle == '(') or (len(First) > spacex+2)):
               Middle=""                                    # not really a middle initial
            First = First[0 : spacex]                       # truncate to actual first Name

      try:
         nnx = nick.loc[(nick['NickName'].str.lower() == First)]
         if(nnx.empty):
            Given=""
         else:
            NickName=nnx.values.flatten().tolist()
            Given=NickName[1]
            Given=Given.lower()
      except KeyError:
         Given=""
      
      # find matching SOS base file record for this member
      # return as list type in mbrow
      lookupbase(First, Last, Middle, Given)

      # find matching Precinct-to-District record for this member
      pnum = str(mrow[0]) + "00"                         # Precinct to Distric Has two trailing zeros on Pct#
      pnum = np.int64(pnum)                              # cast to NumPy INT64 type to match DataFrame typing
      try:
         pcx = pct[pct['PRECINCT'] == pnum]              # find precinct row in cross file
         pctrow = pcx.values.flatten().tolist()          # fetch row into list
         #print(buildcsv(pcthead), buildcsv(pctrow))
      except:
         print("Can't Find Precinct-to-District record for precinct " + pnum)

      #
      # now build output line for expanded members file
      #
      outrow = list(mrow)                                # start with original member row
      outrow.extend(pctrow[1:6])                         # add Senate, Assembly, BOE, Regent and Commisioner districts
      outrow.append(mbrow[14])                           # add data registered to vote
      outrow.append(mbrow[25])                           # add Age
      outrow.extend(mbrow[46:55])                        # add voting propensity caclulations
      outrow.extend(mbrow[26:46])                        # add actual voting method history
      out.write(buildcsv(outrow))                        # write it to output .csv file

   out.close()                                           # close the output file 
     

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
