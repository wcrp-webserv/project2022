#************************************************************************************
#                                base_db_tabless.py                                      
#                                                                                   
#  Input is Processed Secretary of State base.csv file.                             
#                                                                                   
#  Output is an series of extracted files to load mysql tables       
#                                                                                   
#  Output files created:                         
#        base_voter                                                        
#        contact_info                                                  
#        vote_history                                               
#        vote_instance                                                
# *********************************************************************************** 

import pandas as pd
import numpy as np
import sys
import os
import argparse
import csv 



sosfile = "base-db.csv"                       # Secretary of State Data with voting results combined
basefile = "base_voter.csv"                    # output base voter file
contactfile = "contact_info.csv"                # output voter contact file
historyfile = "vote_history.csv"                    
instancefile = "vote_instance.csv"                   
base=""
basevoter_fields = ['StateID', 'CountyID', 'Precinct', 'First_name'] 



###############################################################################
#    main: 
#     
#
#
def main():
   global base  
   parser = argparse.ArgumentParser(description='Process the inputs')
   parser.add_argument("-infile", type=str, help="the name of the file to process")
   parser.add_argument("-dist", type=str, help="ad or sd")
   args = parser.parse_args()
   
   # echo the command line
   print (args.infile) 
   
   mycwd = os.getcwd()
   

############################################################################
# open sosfile base for ourter loop
#
   sos_file = open(args.infile, mode='r', encoding='ISO-8859-1')
   sos_reader = csv.DictReader(sos_file, delimiter=',')
   print ("Processing file: {}".format(args.infile))
   
############################################################################
# open all output files write header
#
   basevoter = open(basefile, 'w', newline='')
   basewriter = csv.DictWriter(basevoter, fieldnames = basevoter_fields)
   basewriter.writeheader()
      
      
###############################################################################
# process base_voter file
#
   for row in sos_reader: 
      i = 0
      string = []
      for field in basevoter_fields:
         string.append({field: str (row[field])})
         i = i+1
      basewriter.writerow(string)

   
      try:
         basewriter.writerow(string)
      except (RuntimeError):
         print('problem occured')
            
   
   ############################################################################
   # close files and clean up
   #  
   sos_file.close
   basevoter.close

            

   exit(0)



###############################################################################
#  Standard boilerplate to call the main() function to begin                   *#                                                                              *
################################################################################

if __name__ == '__main__':
   main()
