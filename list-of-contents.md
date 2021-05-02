Election Management
====================
Programs to create a county voter “base.csv” data file from the Secretary of State data file:
---------------

nvvoter1.pl
-----------
Command syntax:
Perl NVVoter.pl  [-infile <SOSfile> ] [-outfile <Outfile>] [-config <configfile>]
SOSfile  - Secretary of State Vote History File (.VtHst. file downloaded from SOS web site)
outfile – processed  vote history file.  Default name is voterdata.csv  in current directory
configfile – Configuration file.  Default is nvconfig.xlsx  in current directory.
		This file indicates the 20 election cycles to compile votes  from each voter .
This program takes the Secretary of state vote history file which has a single row for every vote and processes it into file that contains a single line for each voter listing whether or not they voted in each of 20 election cycles. The particular cycles are specified in “configfile”.    Note:  After running NVVOTER1.PL  the output file must be sorted by voter ID to become input for NVVOTER2.PL.

nvvoter2.pl
---------------
Command syntax:
Perl NVVoter2.pl [-infile <SOSfile>] [-outfile <outfile>] [-config <configfile>]
    [-statsfile <statsfile>] [-pctfile <precinctfile>] [-emailfile <emailfile>]
SOSfile – Secretary of State Eligible Voter File (.ElgbVtr. file downloaded from SOS web site)
Outfile – final base.csv output file. Default is base.csv in current directory.
configfile – Configuration file.  Default is nvconfig.xlsx  in current directory.
		This file indicates the 20 election cycles to compile votes  from each voter .
statsfile – is the voterdata output file form NVVOTER1.PL  sorted by voter ID. Default is voterdata-s.csv.
precinctfile – output file of vote and registration data by precinct. Default is precinct.csv in current directory.
Emailfile – optional file that has email addresses to add to base.csv.  Default is no file used.

 
Program to split a county base.csv file into simplified and formatted precinct spreadsheets.
base_precinct_xlsx.py
---------------------
Command syntax:
py base_precinct_xlsx.py [-s <basefile>]  [-p <precinct]
basefile – Compiled County Secretary of State data file (output of NVVOTER1.pl/NVVOTER2.pl). 
                  Default is base.csv in current directory.
Precint – Selection option to extract a single precinct.   All precincts are extracted.

Output is one or more formatted .xlsx files each with a simplified view of a single precinct.  The output file(s) have the following naming format:
PCTID_pppp_TOTttt_REPrrr_DEMddd_OTHooo.xlsx
pppp = precinct number
ttt = total registered voters in this precinct
rrr = tot registered republicans in this precinct
ddd = total registered democrats in this precinct
ooo = total registered no part or other party voters in this precinct.
 
Program to extend member file with some information from the county base.csv file.
---------------
memberextend.py
---------------------
Command sytax:
MEMBEREXTEND.PY [-s <SOSfile>] [-p <Pctfile>] [-m <memfile>] [-0 <outfile>]
SOSfile – the county base.csv file. Default is base.csv in current directory.
Pctfile – cross reference file of precinct to district.
Memfile – original Excel member spreadsheet
Outfile – Extended member spreadsheet.  Default is extract.csv in current directory.
Note: Program also has as input the file NickNameList.xls which is used to math full names and nicknames. 

Program to report one or more I360 survey results.
---------------
nvreport.pl
---------------------
Command syntax:
NVReport [-infile <filename.csv>] [ -outfile<filename>]  [-survey <path>]
         [-select param,param,...]
-infile <filename> reports from a single file and must be csv format.
-survey <path> reports from survey files in the specified directory. 
Note: In the absence of either -infile or -survey the current working directory  will be used as if a -survey <cwd> were specified.
-select specifies which files in the survey directory will be selected. 
ADnn - selects files from matching assembly district.
SDnn - selects files from matching senate district.
 pn - selects files from the specified survey phase.
rep - selects file that surveyed republican voters.
dem - selects file that surveyed democrat voters.
oth or othr - selects file that surveyed other party voters.
high - selects file that surveyed high propensity voters.
mod - selects file that surveyed moderate propensity voters.
low - selects file that surveyed low propensity voters.
Note: parameters can be combined in any way. Example:
-select p0,AD27,rep,oth,high,mod\n\n";
-pers specifies that only some respondent persuasion(s) will be compiled. 
              C - selects voters who self select as Conservative.\n";
              MC - selects voters who self select as Moderately Conservative.\n";
              M - selects voters who self select as Moderate.\n";
              MP - selects voters who self select as Moderately Progressive.\n";
              P - selects voters who self select as Progressive.\n";
Note: parameters can be combined in any order. Example:\n";
                    -pers C,M,MC\n\n";
-qfile <filename> text file of question text to substitute for the I360 text.\n";
-outfile <filename> specifies the output report file.  Default is report.txt\n";

Note: Input files come from i360 (More should be written here about how that’s done).

