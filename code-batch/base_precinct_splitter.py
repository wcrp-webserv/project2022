#************************************************************************************
#                                PctExtract.py                                      *
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


Sosfile = "base.csv"                       # Secretary of State Data with voting results combined
outfile = "extract.csv"                    # output extended member file
base=""

#*******************************************************
#                                                      *
#  Routine to get command line arguments (if any)      *
#                                                      *
#*******************************************************
def args(argv):
   global Sosfile
   try:
      opts, args = getopt.getopt(argv,"s",["sosfile="])
   except getopt.GetoptError:
      print('PctExtract.py -s <Sosfile>')
      sys.exit(2)
   for opt, arg in opts:
      if opt == '-h':
         print('PctExtract.py -s <Sosfile>')
         sys.exit()
      elif opt in ("-s", "--sosfile"):
         Sosfile = arg
   print("Input files:")
   temp = '   SOS data file is "' + Sosfile + '"'
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
   baserows=len(base.index)
   #
   # get lists of columnn lbels from teh three input files
   #
   basehead=list(base.columns)                        # get SOS data column labels
   outhead = basehead                                 # get precinct extract file column labels
   #
   #  Create a list of precinct #s in pctlist
   #
   pctlist = []
   for item in base["Precinct"]:
      if item not in pctlist:
         pctlist.append(item)                         # add precinct to list
   pctlist.sort()                                     # sort list in ascending order
   #
   # For each precinct, create an extraction of base.csv items only for that precinct
   #
   for PctNum in pctlist:
      extract = base.loc[base["Precinct"] == PctNum]
      count = len(extract.index)
      print ("Precinct " + str(PctNum) + " has " + str(count) + " Rows")
      #
      #  extract is no a dataframe in same format as base.csv but containing
      #  only those rows that hase the column "Precinct" matching PctNum
      #
      #  open precinct output file with temp name extract.csv
      #
      numRep=0                                           # init counters for this precinct
      numDem=0
      numOth=0
      try:
         out = open(outfile,'w',encoding = 'utf-8')      # open output .csv file and write header line to it.
         out.write(buildcsv(outhead))
      except:
         print("Error opening output file " + outfile + "...aborting\n")
         exit(0)
      for x in range(count):
         outrow = list(extract.iloc[x])                  # get next row of dataframe
         party = outrow[15]
         if (party == "Republican"):
            numRep=numRep + 1
         elif (party == "Democrat"):
            numDem=numDem + 1
         else:
            numOth = numOth + 1

         out.write(buildcsv(outrow))                     # write it to output .csv file
      out.close()                                        # close theextract.csv file
      qualname = "PCTID_"+ str(PctNum) + "_TOT_" + str(count) + "_REP" + str(numRep)
      qualname = qualname + "_DEM" + str(numDem) + "_OTH" + str(numOth) + ".csv"
      os.rename(outfile,qualname)                        # rename extract.csv to actual file name                  
   print("Precinct Files Extracted...exiting")
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
