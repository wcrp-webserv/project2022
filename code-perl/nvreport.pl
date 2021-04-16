
#!/usr/bin/perl 
use strict; 
use warnings;  
use File::Basename;
use Getopt::Long qw(GetOptions);
use Time::Piece;
use Spreadsheet::Read;
use Cwd qw(cwd);
use constant PROGNAME => "NVREPORT - "; 

no warnings "uninitialized";

#
#  Define Files & file handles this program will use
#
my $infile = "";                # -infile on cmd line - overrides and reports onlhy this file
my $printFile = "print.txt";    # Log File Name
my $printFileh;                 # Log File Handle
my $reportFile = "";            # Report File Name
my $reportFileh;                # Report File Handle
my $qfile = "";                 # Optional Question Text Substitution File
my $qfileh;                     # Question Text Substitution file handle
#
#  -Survey directory search, -select sub-select and -pers persuasion qualifier variables
#
my $subdir="";
my $selstring="";
my $perstring ="";
my @persOpts =();                   #Persuasion qualifier Option Array
my $NumPersOpts = 0;                # Number of persuasion qualifier options
my @select = ();
my @surveyfiles = ();
my $nextSurvey = 0;                 # index of next file in surveyfiles to load
my $phase = "";                     # phase of survey
my $dist = "";                      # District (AD or SD) of survey
my $party = "";                     # party (rep, dem oth or othr) of survey
my $likely = "";                    # voting propensity (high, mod, weak) of survey
my @opts = ();                      # option processing array
my @selphase = ();                  # phase selection params
my @seldist = ();                   # district selection params
my @selparty = ();                  # party selection params
my @sellikely = ();                 # voting propensity selection params
my $dir = "";
my $names = "";
my $temp = "";

###################################################################################################################
# >>>>>>>>>>>>>>>>>>>>   Here is where Fixed Assumptions About SpreadSheet are Defined   <<<<<<<<<<<<<<<<<<<<<<<< #
# >>>>>>>>>>>>>>>>>>>>           Change these if Spreadsheet format Changes!!!!          <<<<<<<<<<<<<<<<<<<<<<<< #
#                                                                                                                 #
#  Response Text Strings with special meaning                                                                     #
#                                                                                                                 #
my $Take = "Take Survey";       # This Response says this person TOOK Survey                                      #
my $Refuse = "Refused";         # This Response Says person answered but Refused to TAKE Survey                   #
#                                                                                                                 #
#  Definition of Fixed Columns which allow locating the Question Columns and certain fixed data items             #
#       Note: Column A in Excel = Column 0 here, Column 1 = B etc.                                                #
#                                                                                                                 #
my $VolFirstText = "Volunteer First Name";          # Column Heading for Volunteer First Name                     #
my $VolLastText = "Volunteer Last Name";            # Column Heading for Volunteer Last Name                      #
my $PhoneText = "Phone";                            # Column Heading of Phone Number Called                       #
my $ResponseText = "Household response";            # Column Heading for Contact Response (Busy, No Answer, etc.) #
my $DateText = "Response Date";                     # Column Heading for Contact Date                             #
my $FirstQCol = 15;                                 # Column of First Question in Row = 15 (col P in Excel)       #
my $ColsAfterQ = 3;                                 # Number of Columns Following Questions = 3                   #
#                                                                                                                 #
#  These exact text answers allow sorting answers and also allow detection of an survey responent's self          #
#  definituion of their political persuasion.                                                                     #
#                                                                                                                 #
my @DefAns1 = ("Strongly Disagree","Disagree","No Opinion","Agree","Strongly Agree");                             #
my @DefAns2 = ("Conservative","Moderately Conservative","Moderate","Moderate Progressive","Progressive");         #
#                                                                                                                 #
# >>>>>>>>>>>>>>>>>>>>>>>>>    END OF FIXED ASSUMPTIONS ABOUT SPREADSHEET FORMAT    <<<<<<<<<<<<<<<<<<<<<<<<<<<<< #
###################################################################################################################

#
#  Various global Variables
#
my $helpReq = 0;                # Command Line Option -Help selected Flag
my $j = 0;                      # Outer Loop Counter/index
my $i = 0;                      # inner Loop Counter/index
my $v = 0;                      # Volunteer Table Index
my $q =0;                       # Question processing loop counter/index
my @questionSub;                # array of question text to substitue for header text
my $NumQSub;                    # # of substitute questions
#
#  Init first survey call date to today & last survey call date to time object with a
#  date before any survey calls made.  During spreadsheet loading these will be set
#  to the first date calls were made and the last date calls were made.
#
my $firstdate = Time::Piece->new;   # init to today's date, which is after any call made in this data
my $lastdate = Time::Piece->strptime( "01/01/2020", "%m/%d/%Y" );      # init to before any survey done

#
#   Input SpreadSheet Header & Data Row Arrays
#
my @Headings =();                   # Array of Text Headings for spreadsheet
my @ThisRow =();                    # Data from the Row of spreadsheet currently being processed
my $MaxRows = 0;                    # Number of Rows in current Spreadsheet
my $MaxCols = 0;                    # Number of Columns in current Spreadsheet
my $DataRows = 0;
my $bookdata = 0;                   # reference pointer to loaded spreadsheet

#
#  Indexes to columns we will need during processing.  These
#  are discovered from the header text in the fixed configuration above
#
my $DateCol = -1;                   # Call Date Column Index
my $VolFirstCol = -1;               # Volunteer First Name Column Index
my $VolLastCol = -1;                # Volunteer Last Name Column Index
my $PhoneCol = -1;                  # Phone Number Called Column Index
my $ResponseCol = -1;               # Response Code for Call/Contact Column Index

#
#  Response Types And Counts
#
#  $Numresponse is Number of unique responses
#  @ResponseText = Array of text responses
#  @ResponseCnt = parallel array of Number of time this text response was given
#
my $Numresponse = 0;                # Number of Response Types
my @ResponseText =();               # Array of reponse texts
my @ResponseCnt =();                # Array of Times Each Text Logged
my $Response = "";                  # Current Row's response text

#
#  Volunteer Names and Call Attempt & Result Counts
#
my $NumVolunteer =();               # of different Volunteer Names
my @VolunteerName =();              # Array of Volunteer Names     
my @VolunteerAttempts =();          # Array of # of call attempts by each volunteer
my @VolunteerAnswers =();           # Array of # of answers by Volunteer
my @VolunteerContacts =();          # Array of number of Surveys Taken by Volunteer
my @VolunteerRefuse =();            # Array of survey refusals by Volunteer
my @VolunteerPartial =();           # Array of # of Partial Surveys by volunteer (person quit partway thru)
my $Volunteer = "";                 # Current row's Volunteer Name

#
#   Totals for All Volunteers Combined
#
my $NumAttempts = 0;                # Total Attempts for all volunteers
my $NumAnswers = 0;                 # Total Answers for all volunteers
my $NumContacts = 0;                # Total Surveys Taken for all volunteers
my $NumRefuse = 0;                  # Total Refusals for all volunteers
my $NumPartial =0;                  # Partial Surveys

#
#  Phone Number Tables to find Unique Households Called
#
my $NumPhones = 0;                  # Total Unique Phone Numbers
my @PhoneNumber =();                # List of Phone Numbers Dialed
my $PhoneDuplicates = 0;            # Number of times redialed any number

#
# By Day Of Week Tables
#
my $DowIndex = 0;                   # Day of Week Index (0 to 6)
my @DowText = ("Sun","Mon","Tue","Wed","Thu","Fri","Sat");
my @DowAttempts = (0,0,0,0,0,0,0);  # Unique Call Attempts by Day Of Week
my @DowAnswers = (0,0,0,0,0,0,0);   # Answers by Day Of Week
my @DowContacts = (0,0,0,0,0,0,0);  # Surveys Taken by Day of Week
my @DowRefuse = (0,0,0,0,0,0,0);    # Surveys Refused by Day of Week

#
#  Question Global Tables
#
my $NumQuestions = 0;               # Number of Questions in this spreadsheet (Discovered from SpreadSheet Header Processing)
my @QuestionText = ();              # Table of Question Text (Copied from SpreadSheet Header Row)
my @QuestionNumResponse = ();       # Parallel Table of # of responses for each question
my @QuestionResponse = ();          # Parallel Array of Pointers to Table of Response Text for each question
my @QuestionTally = ();             # Parallel Array of Pointers to Table of Count for each Response Text for each Question
my $rtext = "";                     # question response text temp variable
my $done = 0;                       # question initialization loop exit type flag
my $TotQuestions = 0;               # total questions across all spreadsheets
my @QuestionIndex =();              # list of indexes into Parallel tables for this spreadsheet

#
# >>>>>>>>>>>>>>>>>>>>>  Program Start <<<<<<<<<<<<<<<<<<<<<<<
#
#Open Log file for run time messages and errors
#
open( $printFileh, '>' , "$printFile" )
    or die "Unable to open Log File: $printFile Reason: $!";
#
# Parse any parameters from command line
#
$i=0;
GetOptions(
    'survey=s'   => \$subdir,
    'select=s'   => \$selstring,
    'pers=s'     => \$perstring,
    "infile=s"   => \$infile,
    'outfile=s'  => \$reportFile,
    'qfile=s'    => \$qfile,
    'help!'      => \$helpReq,

    ) or  $i = -1; 

#--------------------------------------------------------
#  help requested or bad option -- give option list
#--------------------------------------------------------
if (($helpReq) or ($i == -1)) {
        print "\$helpReq = $helpReq, \$i = $i\n";
        print "NVReport -infile <filename> -outfile<filename> -survey <path> -select param,param,...\n";
        print "    -infile <filename> reports from a single file.\n";
        print "    -survey <path> reports from survey files in the specified directory. \n";
        print "           Note: In the absence of either -infile or -survey the current working directory\n";
        print "           will be used as if a -survey <cwd> were specified.\n\n";
        print "    -select specifies which files in the survey directory will be selected. \n";
        print "           ADnn - selects files from matching assembly district.\n";
        print "           SDnn - selects files from matching senate district.\n";
        print "             pn - selects files from the specified survey phase.\n";
        print "            rep - selects file that surveyed republican voters.\n";
        print "            dem - selects file that surveyed democrat voters.\n";
        print "            oth or othr - selects file that surveyed other party voters.\n";
        print "           high - selects file that surveyed high propensity voters.\n";
        print "            mod - selects file that surveyed moderate propensity voters.\n";
        print "            low - selects file that surveyed low propensity voters.\n";
        print "              Note: parameters can be combined in any way. Example:\n";
        print "                    -select p0,AD27,rep,oth,high,mod\n\n";
        print "    -pers specifies that only some respondent persuasion(s) will be compiled.  \n";
        print "              C - selects voters who self select as Conservative.\n";
        print "              MC - selects voters who self select as Moderately Conservative.\n";
        print "              M - selects voters who self select as Moderate.\n";
        print "              MP - selects voters who self select as Moderately Progressive.\n";
        print "              P - selects voters who self select as Progressive.\n";
        print "              Note: parameters can be combined in any order. Example:\n";
        print "                    -pers C,M,MC\n\n";
        print "    -qfile <filename> text file of question text to substitute for the I360 text.\n";
        print "    -outfile <filename> specifies the output report file.  Default is report.txt\n";
        exit;
    }


#----------------------------------------------------------
#  Handle -select parameters if specified on command line
#----------------------------------------------------------
if ( $selstring ne "") {
    #
    #  List the options we scanned
    #
    @opts = split(/[,-]/, $selstring);
    $temp = "-select options are: ";
    #
    #  Now stash them into the arrays by their type
    #
    for $i (0 .. @opts-1) {
        $temp = $temp . "\"$opts[$i]\" ";
        my $chr = lc (substr($opts[$i], 0 , 1));
        if ($chr eq "p" ) {
            push (@selphase, lc $opts[$i]);                             # stash phase selector
            next;
        }
        if (($chr eq "a") or ($chr eq "s")) {
            push (@seldist, lc $opts[$i]);                              # stash district selector
            next;
        }
        if (($chr eq "r") or ($chr eq "o") or ($chr eq "d")) {
            if ((lc $opts[$i]) eq "othr") {
                $opts[$i] = "oth";                                      # make "oth" and "othr" synonyms
            }
            push (@selparty, lc $opts[$i]);                             # stash party selector
            next;
        }
        if (($chr eq "h") or ($chr eq "m") or ($chr eq "l")) {
            push (@sellikely, lc $opts[$i]);                            # stash propensity selector;
            next;
        }
        #
        #  We don't recognize this option type
        #
        printLine("Invalid -select parameter \"$opts[$i]\" ignored! \n");
    }
    printLine("$temp \n");
}


#----------------------------------------------------------
#  Handle -pers parameters if specified on command line
#----------------------------------------------------------
if ( $perstring ne "") {
    #
    #  List the options we scanned
    #
    @opts = split(/[,-]/, $perstring);
    $temp = "-pers restrictions are: ";
    #
    #  Now stash them into the arrays by their type
    #
    for $i (0 .. @opts-1) {
        my $chr = uc $opts[$i];
        if ($chr eq "C" ) {
            push (@persOpts, $DefAns2[0]);                             # stash persausion selector
            $NumPersOpts++;
            $temp = $temp . "\"" . $DefAns2[0] . "\" ";
            next;
        }
        if ($chr eq "MC" ) {
            push (@persOpts, $DefAns2[1]);                             # stash persausion selector
            $NumPersOpts++;
            $temp = $temp . "\"" . $DefAns2[1] . "\" ";
            next;
        }
        if ($chr eq "M" ) {
            push (@persOpts, $DefAns2[2]);                             # stash persausion selector
            $NumPersOpts++;
            $temp = $temp . "\"" . $DefAns2[2] . "\" ";
            next;
        }
        if ($chr eq "MP" ) {
            push (@persOpts, $DefAns2[3]);                             # stash persausion selector
            $NumPersOpts++;
            $temp = $temp . "\"" . $DefAns2[3] . "\" ";
            next;
        }
        if ($chr eq "P" ) {
            push (@persOpts, $DefAns2[4]);                             # stash persausion selector
            $NumPersOpts++;
            $temp = $temp . "\"" . $DefAns2[4] . "\" ";
            next;
        }
        #
        #  We don't recognize this persuasion type
        #
        printLine("Invalid -pers parameter \"$opts[$i]\" ignored! \n");
    }
    printLine("$temp \n");
}
#
#--------------------------------------------------------------------
#  Find and Select the Survey Files this program will report
#--------------------------------------------------------------------
#
if ($infile ne "") {
    push (@surveyfiles, $infile);                                   # infile overrides any selection of files
}else{
    $dir = cwd;                                                      # Get our current working directory
    if ($subdir ne "") {
        $dir = $dir . "\/$subdir";                                  # if a Survey Subdirectory specified add it to path
    }
    printLine("Survey Directory = $dir\n");

    opendir my $dh, $dir
        or die "Could not open '$dir' for reading: $! \n";          # open the survey directory
    while($names = readdir $dh) {                                   # read in file names
        $temp = lc $names;                                          # force lower case
        if ( index($temp, ".xlsx") == -1) {
            next;
        } 
        if (substr($temp, 0, 5) eq "i-20p" ) {                      # if name starts with i-20p it's a survey
            @opts = split (/-/, $temp);                             # break out lower cased file name into segments
            $temp = substr($opts[1], 2, length($opts[1])-2);        # extract phase of survey
            $phase = $temp;                                         # store phase
            $dist = $opts[2];                                       # store district
            if ($opts[3] eq "othr") {
                $opts[3] = "oth";                                   # make "othr" synonym for "oth"
            }
            $party = $opts[3];                                      # story party
            $opts[4]  =~ s/.xlsx//;                                 # strip traling file extension
            $likely = $opts[4];                                     # store voting likelyhood
            $temp = qualify_file();                                 # screen any seletion criteria
            if ($temp == 0) {
                push (@surveyfiles, $names);                        # This file is selected add to survey file name list
            }
        }
    }
    closedir $dh;                                                   # done, close directory
}
#
#  The array @surveyfiles is now a list of zero or more survey files to report
#
if (scalar @surveyfiles == 0) {
    printLine ("No input files found -- nothing to report...\n");
    exit;
}

if ( $reportFile eq "")
    {
    #
    #  No report file name on command line, form default report file name
    #
    $reportFile = "report.txt";
    }

# 
#Open report file for the Report itself
#
open( $reportFileh, '>' , "$reportFile" )
    or die "Unable to open Report File: $reportFile Reason: $!";
printLine("Reporting to file $reportFile...\n");

#
# Open & Load -infile spreadsheet into array pointed to by $bookdata
# 
$infile =$surveyfiles[$nextSurvey];                             # get name of 1st spreadsheet
if ($subdir ne "") {
    $infile = $subdir . "/" . $infile;                          # prepend subdirectory
}
$nextSurvey++;                                                  # say it's been used
printLine("Loading spreadsheet $infile ...\n");
$bookdata = Spreadsheet::Read->new($infile)
     or die "Unable to open input Spreadsheet File: $infile Reason: $!";
init_spreadsheet();
#
#  Initialize Question Discovery and Tally Variables for 1st spreadsheet
#
$NumQuestions = $MaxCols - $FirstQCol - $ColsAfterQ;            # Calculate The Number of Questions in this spreadsheet
$TotQuestions = $NumQuestions;                                  # first spreadsheet, this is also all the questions so far
@QuestionIndex =();                                             # reset question index array
for $q (0 .. ($NumQuestions-1))
{
    push(@QuestionIndex, $q);                                   # for first file, index = 1, 2, ....
    push(@QuestionText, $Headings[$q+$FirstQCol]);              # Build Question Text Table
    push(@QuestionResponse, 0);                                 # create correct number of pointer entries
    push(@QuestionTally, 0);                                    # for reference pointer arrays
    push(@QuestionNumResponse, 0);                              # No Responses so far
    my @QResponse = ();                                         # Make empty element arrays
    my @QTally = ();                                            # For Response Text and Tally Counters
    $QuestionResponse[$q] = \@QResponse;                        # Add Pointers to them to dynamic Response Text arrays
    $QuestionTally[$q] = \@QTally;
}
printLine("... Initialized for $NumQuestions Questions ...\n");
#
# >>>>>>>>>>>>>>>>  Basic Initialization done for 1st file .... <<<<<<<<<<<<<<<<<<<<<<<<
#
NEXTFILE:
#
# Cycle Through spreadsheet rows and build the various arrays needed for reporting
#
for $j (2 .. $MaxRows)
{

    @ThisRow = Spreadsheet::Read::row($bookdata->[1], $j);          # Fetch next row of spreadsheet

    #---------------------------------------------------------------------
    #  If there are any -pers persuasion qualifier options, see if this row qualifies
    #
    if ($NumPersOpts > 0) {
        for $q (0 .. ($NumQuestions-1))
        {
            for my $ii (0 .. ($NumPersOpts-1))
            {
                if ($ThisRow[$FirstQCol+$q] eq $persOpts[$ii])
                {
                    goto ROWQUALIFIES;                      # this row qualifies, process it
                }
            }
        }
        goto SKIPROW;                                       # skip processing this row
    }
    ROWQUALIFIES:
    #
    #----------------------------------------------------------------------
    #
    #  Process this row of the input spreadsheet
    #
    $NumAttempts++;                      # Every Row is an attempt

    #
    #  Get Index 0-6 for the Day of the Week this Call Was Made
    #
    my ( @date, $yy, $mm, $dd, $TextDate, $cdate, $mx, $dx, $yx);
    #
    #  Note:  Date in spreadsheet may have either "-" or "/" delimiter
    #         It may also be in either yyyy-mm-dd or mm-dd-yyyy order
    #         The following is to detect its format and convert
    #         it to a known mm/dd/yyyy format, then to a Time::Piece object,
    #         and from that object get the day of the week index 0-6.
    #
    @date = split( /[-,\/ ]/, $ThisRow[$DateCol], -1 );         # fetch yyyy-mm-dd format date from row and split into yyyy, mm, dd in @date
    if ((length $date[0]) == 4)
    {
        $yx=0;              # format is yyyy-mm-dd
        $mx=1;
        $dx=2
    }else{
        $mx=0;              # Format = mm-dd-yyyy
        $dx=1;
        $yx=2;
    }
    $mm = sprintf( "%02d", $date[$mx] );                        # create mm, dd, yyyy as separate strings
    $dd = sprintf( "%02d", $date[$dx] );
    $yy = sprintf( "%02d", substr( $date[$yx], 0, 4 ) );
    $TextDate = "$mm/$dd/$yy";                                  # assemble them into mm/dd/yyyy single string
    $cdate = Time::Piece->strptime( $TextDate, "%m/%d/%Y" );    # build Time::Piece object of midnight on that date
    $DowIndex = $cdate->day_of_week;                            # get the Day of the Week index (0=Sunday, 1=Monday, etc.,)
    $DowAttempts[$DowIndex]++;                                  # log attempts by day of week
    #
    # See if this call date may be the first or last date of survey interval
    #
    if ($cdate < $firstdate)
    {
        $firstdate = $cdate;                                    # earlier than current first date
    }
    if ($cdate > $lastdate)
    {
        $lastdate = $cdate;                                     # later than current last date
    }

    #
    #  See if Unique Phone Number, if so add to table.  If Not, count duplicates
    #
    if ($NumPhones > 0)
    {
        for $i (0 .. $NumPhones-1)
        {
            if($PhoneNumber[$i] eq "")
            {
                goto EXITPHONE;                                # don't add blank to table
            }
            if($PhoneNumber[$i] eq $ThisRow[$PhoneCol])
            {
                $PhoneDuplicates++;                            # this is a duplicate Phone Number
                goto EXITPHONE;                                # don't add to table
            }
        }
        push(@PhoneNumber, $ThisRow[$PhoneCol]);                       # Unique Phone Number so far, add to table
        $NumPhones++;                                          # Count Unique Numbers Called
    }else{
        push(@PhoneNumber, $ThisRow[$PhoneCol]);                       # First Phone Nubmer Called, add to table
        $NumPhones++;                                          # Count Unique Numbers Called
    }
    EXITPHONE:

    #
    # Build @ResponseText array and log count for each text response in $ResponseCnt array
    #
    $Response = $ThisRow[$ResponseCol];            # fetch response text
    if ($Numresponse > 0)
    {
        for $i (0 .. ($Numresponse - 1))
        {
            if ($Response eq $ResponseText[$i])
            {
                $ResponseCnt[$i]++;                     # Already present, count one more hit
                goto EXITRESPONSE;
            }
        }
        AddResponse();
    }else{
        AddResponse();   
    }
    EXITRESPONSE:

    # Build Volunteer Name Array and Counts by Volunteer
    $Volunteer = join(" ",$ThisRow[$VolFirstCol], $ThisRow[$VolLastCol]);       # Fetch and Form Full First + Last Name
    if ($NumVolunteer > 0)
    {
        for $i (0 .. ($NumVolunteer - 1))           # See if Name Has already been found
        {
            if ($Volunteer eq $VolunteerName[$i])
            {
                $VolunteerAttempts[$i]++;           # Name Already in Tables, count one more attempt
                $v = $i;                            # Set index of this volunteer
                goto EXITVOLUNTEER;
            }
        }
        AddVolunteer();                             # Add new name to Tables
        $v = $NumVolunteer-1;                       # set index to newly added Volunteer
    }else{
        AddVolunteer();                             # Add First Name to Tables
        $v = $NumVolunteer-1;                       # set index to newly added Volunteer
    }
    EXITVOLUNTEER:
    #
    #  Whether new, or repeat name, $v is index of this volunteer's table entries
    #
    #  Process and Log the "by volunteer" data, the total data, and the "by day of week called" data
    #
    if ($ThisRow[$ResponseCol] eq $Take)            # See if this person took survey
    {
        $VolunteerContacts[$v]++;                   # Survey = Answer & Taken
        $VolunteerAnswers[$v]++;
        $NumAnswers++;                              # Add to total Answers
        $NumContacts++;                             # add to total Surveys Taken
        $DowContacts[$DowIndex]++;                  # also by day of week
        $DowAnswers[$DowIndex]++;                   # Answers too
        if (doquestions($j) == 1)
        {
            $NumPartial++;
            $VolunteerPartial[$v]++;                # say this was a partial survey
        }
    }
    if ($ThisRow[$ResponseCol] eq $Refuse)          # See if this person Refused to take survey
    {
        $VolunteerAnswers[$v]++;                    # Refuse = Answer + Refuse
        $VolunteerRefuse[$v]++;
        $NumAnswers++;                              # Add to total Answers
        $NumRefuse++;                               # and total Refuses
        $DowRefuse[$DowIndex]++;                    # also by day of week
        $DowAnswers[$DowIndex]++;                   # Answers too
    }

    SKIPROW:
}
#
# >>>>> End Processing of Spreadsheet Data into Tables and Counters <<<<<
#
#       **************************************************************************
# >>>>> See if there is another spreadsheet to load, or if it's time to do reports <<<<<<
#       **************************************************************************
#
if ($nextSurvey < scalar @surveyfiles) {
    open_next_survey();
    goto NEXTFILE;                                                  # load/process next survey file
}
#
#  See if any question text should be substituted from an optional question text file
#
my $numsub = 0;
if (open($qfileh, '<', $qfile)) {
    while (my $row = <$qfileh>)
    {
        #
        #  find matching question and substitute text
        #  consider question a match if first 20 charactrers match
        #
        chomp $row;
        $i=length($row);
        if ($i < 3) {
            next;                       # skip blank lines
        }
        if ($i > 50) {
            $i = 50;                    # check max of 50 characters for match
        }
        for $q (0 .. ($TotQuestions-1))
        {
            if ( (lc substr($row,0,$i)) eq (lc substr($QuestionText[$q],0,$i)) ) {
                $QuestionText[$q] = $row;   # substitute this text
                $numsub++;
                last;
            }
        }
    }
} else {
    printLine ("Could not open file question '$qfileh' reason: $! \n");
    printLine ("Continuing with survey file queston text...\n");
}
if ($numsub > 0) {
    printLine ("$numsub questions had text substituted...\n");
}
#
#  >>>>>>>>>>>>> All Survey Files are loaded, do the reporting <<<<<<<<<<<<<<<<<<<<
#
printLine("... Generating Reports to file $reportFile ...\n");
&ReportAttempts();
&ReportCallStatus();
&ReportByVolunteer();
&ReportByDayOfWeek();
&ReportQuestions();
#
#   End of Program, close files and exit
#
EXIT:
close($printFileh);
close($reportFileh);
exit;
#
# >>>>>>>>>>>>>>>>>>>>>>>   End Main Program <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
#
# >>>>>>>>>>>>>>>>>>>>>>>  Begin SubRoutines <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

#----------------------------------------------------------------
#   Generate report header plus the Call Attempt/Result Report
#----------------------------------------------------------------
sub ReportAttempts {
    printLine("Survey Covers $NumAttempts Calls with $NumAnswers Answers\n");
    #
    #  Calculate and/or format data items
    #
    my $pfirst = $firstdate->mdy;           # get survey start date mm-dd-yyyy
    my $plast = $lastdate->mdy;             # get survey end date mm-dd-yyyy
    my $attempts = sprintf( "%4d", $NumAttempts );
    my $Contacts = sprintf( "%4d", $NumContacts );
    my $partial = sprintf( "%4d", $NumPartial );
    my $surveyPct = FourPct($NumContacts/$NumAttempts*100);
    my $redial = sprintf( "%4d", $PhoneDuplicates );
    my $refuses = sprintf( "%4d", $NumRefuse );
    my $refusePct = FourPct($NumRefuse/$NumAnswers*100);
    my $unique = sprintf( "%4d", ($NumAttempts-$PhoneDuplicates));
    my $Answers = sprintf( "%4d", $NumAnswers );
    my $AnswerPct =FourPct($NumAnswers/$NumAttempts*100);
    #
    # Print the report lines
    #
    if (scalar @surveyfiles == 1) {
        print $reportFileh("                 Reporting From File $surveyfiles[0]\n");
    }else{
        print $reportFileh("                 Reporting From Combined Files:\n");
        for $i (0 .. @surveyfiles-1) {
            print $reportFileh("                 $surveyfiles[$i]\n");
        }
    }
    if ($NumPersOpts > 0){
        #
        #  print that  this is conditional partial report of selected persuasion respondents
        #
        print $reportFileh("\n   Reporting only Persuasion(s) " . join(',', @persOpts) . "\n\n");
    }
    print $reportFileh("\n\n                 Survey Interval from $pfirst to $plast\n\n\n");
    print $reportFileh("                      Total Survey Call Effectiveness\n");
    print $reportFileh("                      -------------------------------\n");
    print $reportFileh("Call Attempts: $attempts       Total Answers: $Answers               Answer%: $AnswerPct%\n");
    print $reportFileh("      Redials: $redial     Survey Contacts: $Contacts (Part $partial)   Survey%: $surveyPct%\n");
    print $reportFileh(" Total Unique: $unique     Survey Refusals: $refuses               Refuse%: $refusePct%\n\n\n");
    return;
}

#---------------------------------------------------------------
#  Generate Call Response Summary Report
#---------------------------------------------------------------
sub ReportCallStatus() {
    printLine("Total Call Response Types: $Numresponse\n");
    print $reportFileh("                         Incomplete Call Breakdown\n");
    print $reportFileh("                         -------------------------\n");
    for $i (0 .. ($Numresponse - 2))
    {
        if ($ResponseText[$i] eq $Take)
        {
            next;
        }
        if ($ResponseText[$i] eq $Refuse)
        {
            next;
        }
        print $reportFileh("$ResponseText[$i] $ResponseCnt[$i], ");
    }
    my $last = $Numresponse-1;
    print $reportFileh("$ResponseText[$last] $ResponseCnt[$last]\n\n\n");
    return;
}

#---------------------------------------------------------------
#  Generate the by Volunteer Report
#---------------------------------------------------------------
sub ReportByVolunteer {
    printLine("Total Volunteer Names: $NumVolunteer\n");
    my $name = "";                  # Defin local working variables
    my $attempts = "";
    my $Answers = "";
    my $Contacts = "";
    my $refuse = "";
    my $completePct = "";
    my $AnswerPct = "";
    print $reportFileh("                         Performance by Volunteer\n");
    print $reportFileh("                         ------------------------\n\n");
    print $reportFileh("     Volunteer Name   Attempts  Answers   Contacts  Refuse  Completion%  Answer%\n");
    print $reportFileh("     --------------   -------   -------   --------  ------  -----------  --------\n");
    for $i (0 .. ($NumVolunteer - 1))
    {
        $name = sprintf("%19s", $VolunteerName[$i] );               # format items to print
        $attempts = sprintf( "%4d", $VolunteerAttempts[$i] );
        $Answers = sprintf( "%4d", $VolunteerAnswers[$i] );
        $Contacts = sprintf( "%4d", $VolunteerContacts[$i] );
        $refuse = sprintf( "%4d", $VolunteerRefuse[$i] );
        $completePct = FourPct($VolunteerContacts[$i]/$VolunteerAttempts[$i]*100);
        $AnswerPct = FourPct($VolunteerAnswers[$i]/$VolunteerAttempts[$i]*100);
        print $reportFileh("$name     $attempts      $Answers       $Contacts    $refuse      $completePct%       $AnswerPct%\n");
    }
    print $reportFileh("\n");
    return;
}

#--------------------------------------------------------------
# Generate the Day Of Week Report
#--------------------------------------------------------------
sub ReportByDayOfWeek {
    my $sun;                        # declare local formatting string variables
    my $mon;
    my $tue;
    my $wed;
    my $thu;
    my $fri;
    my $sat;
    print $reportFileh("                        Breakdown By Day Of Week\n");
    print $reportFileh("                        ------------------------\n");
    print $reportFileh("             Sun       Mon       Tue       Wed       Thu       Fri      Sat\n");
    print $reportFileh("           -------   -------   -------   -------   -------   -------  -------\n");
    $sun = sprintf("%4d", $DowAttempts[0]);
    $mon = sprintf("%4d", $DowAttempts[1]);
    $tue = sprintf("%4d", $DowAttempts[2]);
    $wed = sprintf("%4d", $DowAttempts[3]);
    $thu = sprintf("%4d", $DowAttempts[4]);
    $fri = sprintf("%4d", $DowAttempts[5]);
    $sat = sprintf("%4d", $DowAttempts[6]);
    print $reportFileh(" Attempts   $sun      $mon      $tue      $wed      $thu      $fri     $sat\n");
    $sun = sprintf("%4d", $DowAnswers[0]);
    $mon = sprintf("%4d", $DowAnswers[1]);
    $tue = sprintf("%4d", $DowAnswers[2]);
    $wed = sprintf("%4d", $DowAnswers[3]);
    $thu = sprintf("%4d", $DowAnswers[4]);
    $fri = sprintf("%4d", $DowAnswers[5]);
    $sat = sprintf("%4d", $DowAnswers[6]);
    print $reportFileh("  Answers   $sun      $mon      $tue      $wed      $thu      $fri     $sat\n");
    $sun = sprintf("%4d", $DowContacts[0]);
    $mon = sprintf("%4d", $DowContacts[1]);
    $tue = sprintf("%4d", $DowContacts[2]);
    $wed = sprintf("%4d", $DowContacts[3]);
    $thu = sprintf("%4d", $DowContacts[4]);
    $fri = sprintf("%4d", $DowContacts[5]);
    $sat = sprintf("%4d", $DowContacts[6]);
    print $reportFileh(" Contacts   $sun      $mon      $tue      $wed      $thu      $fri     $sat\n");
    $sun = sprintf("%4d", $DowRefuse[0]);
    $mon = sprintf("%4d", $DowRefuse[1]);
    $tue = sprintf("%4d", $DowRefuse[2]);
    $wed = sprintf("%4d", $DowRefuse[3]);
    $thu = sprintf("%4d", $DowRefuse[4]);
    $fri = sprintf("%4d", $DowRefuse[5]);
    $sat = sprintf("%4d", $DowRefuse[6]);
    print $reportFileh("   Refuse   $sun      $mon      $tue      $wed      $thu      $fri     $sat\n");
    for my $x (0 .. 6)
    {
        if ($DowAttempts[$x] == 0)
        {
            $DowAttempts[$x] = 0.1;             # Avoid divide by zero
        }
    }
    $sun = FourPct($DowAnswers[0]/$DowAttempts[0]*100);        # Format Answer %
    $mon = FourPct($DowAnswers[1]/$DowAttempts[1]*100);
    $tue = FourPct($DowAnswers[2]/$DowAttempts[2]*100);
    $wed = FourPct($DowAnswers[3]/$DowAttempts[3]*100);
    $thu = FourPct($DowAnswers[4]/$DowAttempts[4]*100);
    $fri = FourPct($DowAnswers[5]/$DowAttempts[5]*100);
    $sat = FourPct($DowAnswers[6]/$DowAttempts[6]*100);
    print $reportFileh("  Answer%   $sun%     $mon%     $tue%     $wed%     $thu%     $fri%    $sat%\n");
    $sun = FourPct($DowContacts[0]/$DowAttempts[0]*100);        # Format Complete %
    $mon = FourPct($DowContacts[1]/$DowAttempts[1]*100);
    $tue = FourPct($DowContacts[2]/$DowAttempts[2]*100);
    $wed = FourPct($DowContacts[3]/$DowAttempts[3]*100);
    $thu = FourPct($DowContacts[4]/$DowAttempts[4]*100);
    $fri = FourPct($DowContacts[5]/$DowAttempts[5]*100);
    $sat = FourPct($DowContacts[6]/$DowAttempts[6]*100);
    print $reportFileh("Complete%   $sun%     $mon%     $tue%     $wed%     $thu%     $fri%    $sat%\n");
    for my $x (0 .. 6)
    {
        if ($DowAnswers[$x] == 0)
        {
            $DowAnswers[$x] = 0.1;             # Avoid divide by zero
        }
    }
    $sun = FourPct($DowRefuse[0]/$DowAnswers[0]*100);        # Format Answer %
    $mon = FourPct($DowRefuse[1]/$DowAnswers[1]*100);
    $tue = FourPct($DowRefuse[2]/$DowAnswers[2]*100);
    $wed = FourPct($DowRefuse[3]/$DowAnswers[3]*100);
    $thu = FourPct($DowRefuse[4]/$DowAnswers[4]*100);
    $fri = FourPct($DowRefuse[5]/$DowAnswers[5]*100);
    $sat = FourPct($DowRefuse[6]/$DowAnswers[6]*100);
    print $reportFileh("  Refuse%   $sun%     $mon%     $tue%     $wed%     $thu%     $fri%    $sat%\n");
    print $reportFileh("\n\n");
}

#--------------------------------------------------------------
#  Generate the Question & Answer report
#--------------------------------------------------------------
#
#  Two answer sets that if we find them, we want in a specified order
#
#
sub ReportQuestions {
    my @OrderResp = ();
    my @OrderTally = ();
    my $reorder;
    my $x;
    my $TotalTally = 0;
    if (scalar @surveyfiles == 1) {
        printLine("Survey has $TotQuestions Questions.\n");
    }else{
        printLine("Combined the Surveys have $TotQuestions Total Questions.\n");
    }


    for $q (0 .. ($TotQuestions - 1))
    {
        $reorder = 0;                                               # assume answer don't need to be reordered
        @OrderResp = (" ") x ($QuestionNumResponse[$q] - 1);        # init report arrays to proper length
        @OrderTally = (0)  x ($QuestionNumResponse[$q] - 1);        # with 0 Tallys
        #
        $QuestionText[$q] =~ s/\n/\n   /g;                         # Indent any newlines 3 spaces
        print $reportFileh("Q: $QuestionText[$q]\n");   # Print Question Text

        #
        #  Add up total responses to this question for later percentage calculations
        #
        $TotalTally = 0;                                            # Init Total Tally
        for $i (0 .. ($QuestionNumResponse[$q] - 1))
        {
            $TotalTally = $TotalTally + $QuestionTally[$q]->[$i];   # add all answer counts for this question
        }
        #
        #  See if this is a recognized response list
        #
        for $i (0 .. ($QuestionNumResponse[$q] - 1))
        {
            for $x (0 .. 4)
            {
                if ($QuestionResponse[$q]->[$i] eq $DefAns1[$x])            
                {
                    $reorder = 1;               # indicate needs reordering against DefAns1
                    goto REORDER;
                }
                if ($QuestionResponse[$q]->[$i] eq $DefAns2[$x])
                {
                    $reorder = 2;               # indicate needs reordering against DefAns2
                    goto REORDER;
                }
            }
        }
        #
        #  Isn't a response list we know, order it as discovered
        #
        for $i (0 .. ($QuestionNumResponse[$q] - 1))
        {
            $OrderResp[$i] = $QuestionResponse[$q]->[$i];
            $OrderTally[$i] = $QuestionTally[$q]->[$i];
        }
        goto REPORT;
        #
        #  This is at least partly a recognized list, Answers must be reordered
        #
        REORDER:
        my $addx = 5;                                           #add 1st unknown here
        if ($reorder == 1)
        {
            for $x (0 .. 4)
            {
                $OrderResp[$x] = $DefAns1[$x];                   # start with defined text order
            }
            for $i (0 .. ($QuestionNumResponse[$q] - 1))
            {
                for $x (0 .. 4)
                {
                    if (lc($QuestionResponse[$q]->[$i]) eq lc($DefAns1[$x]))     # this is one of the defined responses
                    {
                        $OrderTally[$x] = $QuestionTally[$q]->[$i];  #move Tally count to proper slot
                        goto FOUND1;
                    }
                }
                #
                #  Not one of the defined answers, put at end
                #
                $OrderResp[$addx] = $QuestionResponse[$q]->[$i];    # add to next unknown response slot
                $OrderTally[$addx] = $QuestionTally[$q]->[$i];
                $addx++;                                            # say this slot used
                FOUND1:
            }
        }else{
            for $x (0 .. 4)
            {
                $OrderResp[$x] = $DefAns2[$x];                   # start with defined text order
            }
            for $i (0 .. ($QuestionNumResponse[$q] - 1))
            {
                for $x (0 .. 4)
                {
                    if (lc($QuestionResponse[$q]->[$i]) eq lc($DefAns2[$x]))     # this is one of the defined responses
                    {
                        $OrderTally[$x] = $QuestionTally[$q]->[$i];  #move Tally count to proper slot
                        goto FOUND2;
                    }
                }
                #
                #  Not one of the defined answers, put at end
                #
                $OrderResp[$addx] = $QuestionResponse[$q]->[$i];    # add to next unknown response slot
                $OrderTally[$addx] = $QuestionTally[$q]->[$i];
                $addx++;                                            # say this slot used
                FOUND2:
            }
        }
        REPORT:
        #
        #  Now Report the reordered Array
        #
        my $Slen = 0;                                           # Response Text Character Count
        my $TCnt = "";                                          # Formatted Tally Count
        my $RText = "";                                         # Formatted Response Text
        my $TPct = "";                                          # Formatted Percentage text
        my $BText = "";                                         # Bar Graph Text
        my @Words = ();                                         # Multiword array for multi-line response formatting.
        my $WdCnt = 0;                                          # # of words in multi word array
        my $Mx = 0;                                             # Multiword array index
        my $x = 0;                                              # local loop variable

        for $i (0 .. ($QuestionNumResponse[$q] - 1))
        {
            $Slen = length $OrderResp[$i];                       # Get Reponse Text
            $TCnt = sprintf("%4d", $OrderTally[$i]);             # Get and Format Response Tally Count
            if ($Slen < 26)
            {
                #
                #  Response fits on single line
                #  Build string to print
                #
                $RText = $OrderResp[$i] . (".") x (28-$Slen);
                $Mx = -1;                                       # Indicate not multi-line response
            }else{
                #
                #  Response Text is multi-line
                #
                #  Array of words and then build first line of response
                #
                @Words = split /\s+/, $OrderResp[$i]; ;         # Split Response Text into Words
                $WdCnt = scalar @Words;                         # and count the words that result
                $RText = $Words[0];                             # Init Respose String to 1st word
                $Slen = length $RText;                          # init String length
                $Mx = 1;                                        # Point to next word
                while ($Slen < 26)
                {
                    if (($Slen + 1 + length $Words[$Mx]) < 26)
                    {
                        #
                        #  Room to add next word preceeded by a space
                        #
                        $RText = $RText . " " . $Words[$Mx];    # add space and next word
                        $Slen = length $RText;                  # update string length
                        $Mx++;                                  # point to next word
                    }else{
                        #
                        #  Can't add next word, it will be first word of next line
                        #  finish out this line in $Rtext
                        #
                        $RText = $RText . (".") x (28-$Slen);
                        $Slen = 30;                             # assure exit from while loop
                        #
                        #  Note $Mx is now index of the 1st word in the next line
                        #
                    }
                }
            }
            #
            #  Calculate the percantage for this response
            #
            $TPct = FourPct($TCnt/$TotalTally*100);             # Formatted and rounded % to print
            $Slen = int((($TCnt/$TotalTally*100)/3)+ 0.5);      # get 1/3 of percent (rounded up) = # bar chars
            if ($Slen == 0)
            {
                $Slen = 1;                                       # minimum 1 bar char
            }
            $BText = ("■") x $Slen;                                # Build Graph Bar Text
            print $reportFileh("       " . $RText . $TCnt . "   " . $TPct . "%  " . $BText . "\n");
            #
            #  If multiline response, print rest of the line(s)
            #
            NEXTLINE:
            if ( ($Mx > 0) && ($Mx < $WdCnt) )
            {
                $RText = $Words[$Mx];                           # Init Respose String to next word
                $Slen = length $RText;                          # init String length
                $Mx++;                                          # point to next word
                while ( ($Slen < 28) && ($Mx < $WdCnt) )
                {
                    if ( ( $Slen + 1 + length $Words[$Mx]) <= 28 )
                    {
                        #
                        #  Room to add next word preceeded by a space
                        #
                        $RText = $RText . " " . $Words[$Mx];    # add space and next word
                        $Slen = length $RText;                  # update string length
                        $Mx++;                                  # point to next word
                    }else{
                        $Slen = 30;                             # all that fits on line, exit while loop
                    }
                }
                print $reportFileh("       " . $RText . "\n");  # print this line
            }
            if ( ($Mx > 0) && ($Mx < $WdCnt) )
            {
                goto NEXTLINE;                                  # Another Line to Responde Code
            }
        }
        print $reportFileh("\n");
    }
}

#-------------------------------------------------------------
#  Take input % fraction and format to a string with
#  1 decimal point that is 4 characters long.
#-------------------------------------------------------------
sub FourPct {
    my ($fraction) = @_;
    my $cstring = sprintf( "%.1f", ($fraction) );
    if (length $cstring == 3)
    {
        $cstring = " " . $cstring;
    }
    return $cstring;
}

#-------------------------------------------------------------
#  Add Response text to Response table
#-------------------------------------------------------------
sub AddResponse {
    push @ResponseText, $Response;                  # New Response Text Add With 1 Hit count
    push @ResponseCnt, 1;
    $Numresponse++;
}

#-------------------------------------------------------------
# Add Volunteer Name to Volunteer Table
#-------------------------------------------------------------
sub AddVolunteer {
    push @VolunteerName, $Volunteer;                # New Name, Add to name array
    push @VolunteerAttempts, 1;                     # with 1 attempt
    push @VolunteerAnswers, 0;                     # init rest of numbers to 0
    push @VolunteerContacts, 0;
    push @VolunteerPartial, 0;
    push @VolunteerRefuse, 0;
    $NumVolunteer++;
}

#------------------------------------------------------------
#  Load/Process Questions for this row
#------------------------------------------------------------
sub doquestions {
    my ($row) = @_;
    my $rcode = 0;
    my $z = 0;
    for $q (0 .. ($NumQuestions-1))
    {
        $rtext = $ThisRow[$FirstQCol+$q];               # Fetch This question Response
        if ((length $rtext) == 0)
        {
            $rcode = 1;                                 # flag Partial Survey
            next;                                       # skip this question
        }
        $z=$QuestionIndex[$q];                          # get the question index into the parallel table array
        if ($QuestionNumResponse[$z] == 0)
        {
            push(@{$QuestionResponse[$z]} , $rtext);    # Add This Response for This question
            push(@{$QuestionTally[$z]} , 1);            # It has Happened once
            $QuestionNumResponse[$z] = 1;               # this question now has 1 response
        }else{
            $done=0;                                    # init loop exit type flag
            for $i (0 .. ($QuestionNumResponse[$z]-1))
            {
                if ($QuestionResponse[$z]->[$i] eq $rtext)
                {
                    $QuestionTally[$z]->[$i]++;         # Another hit for this response
                    $done = 1;                          # done for this question
                }
            }
            if ($done == 0)
            {
            push(@{$QuestionResponse[$z]} , $rtext);    # Add This Response for This question
            push(@{$QuestionTally[$z]} , 1);            # It has Happened once
            $QuestionNumResponse[$z]++;                 # this question now has another response
            }
        }
    }
    return $rcode;                                      #all questions done
}
#----------------------------------------------------------------------
#  Open and Load 2nd to nth survey file for -survey option
#----------------------------------------------------------------------
sub open_next_survey {
    #
    # Open & Load next -survey spreadsheet into array pointed to by $bookdata
    #
    my $new = 0;                                                    # assum no new questions
    $infile =$surveyfiles[$nextSurvey];                             # get name of 1st spreadsheet
    if ($subdir ne "") {
        $infile = $subdir . "/" . $infile;                          # prepend subdirectory
    }
    $nextSurvey++;                                                  # say it's been used
    printLine("Loading spreadsheet $infile ...\n");
    $bookdata = Spreadsheet::Read->new($infile)
        or die "Unable to open input Spreadsheet File: $infile Reason: $!";
    init_spreadsheet();                                             # get basic stuff and verify this is survey spreadsheet
    #
    #  Initialize Question Discovery and Tally Variables for 1st spreadsheet
    #
    $NumQuestions = $MaxCols - $FirstQCol - $ColsAfterQ;            # Calculate The Number of Questions in this spreadsheet
    @QuestionIndex =();                                             # reset question index array
    for $i (0 .. $NumQuestions-1) {
        push (@QuestionIndex, -1);                                  # init QuestionIndex array with -1 in each slot
    }
    for $i (0 .. $NumQuestions-1) {                                 # process these questions to see if new or not
        $temp = -1;                                                 # init as new question
        for $q (0 .. $TotQuestions-1) {
            if ($Headings[$i+$FirstQCol] eq $QuestionText[$q]) {
                $QuestionIndex[$i] = $q;                            # repeat question, point to it in parallel arrays
                $temp=0;                                            # flag repeat question
            }
        }
        if ($temp == -1) {
            $new++;                                                 # count new questions
            $QuestionIndex[$i] = $TotQuestions;                     # index for this question
            $TotQuestions++;                                        # one mor total question
            push(@QuestionText, $Headings[$i+$FirstQCol]);          # Build Question Text Table
            my @QResponse = ();                                     # Make empty element arrays
            my @QTally = ();                                        # For Response Text and Tally Counters
            push(@QuestionResponse, \@QResponse);                   # Add Response array Reference
            push(@QuestionTally, \@QTally);                         # Add Tally Arry reference
            push(@QuestionNumResponse, 0);                          # No Responses so far
        }
    }
    for $i (0 .. $NumQuestions-1) {
        if ($QuestionIndex[$i] == -1) {                             # safety check
            print "Internal Failure in 2nd to nth Survey File load ... aborting\n";
            exit;
        }
    }
    if ($new > 0) {
        printLine("... Initialized for $NumQuestions Questions ($new new) ...\n");
    }else{
        printLine("... Initialized for $NumQuestions Questions ...\n");
    }
    return;
}

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
sub init_spreadsheet {
    $MaxRows = $bookdata->[1]{maxrow};                  # Save out Number of Rows in Spreadsheet
    $MaxCols = $bookdata->[1]{maxcol};                  # Save out Number of Columns in Spreadsheet
    $DataRows = $MaxRows-1;
    printLine("... Input SpreadSheet Loaded: $DataRows Data Rows of $MaxCols columns each.\n"); 
    #
    # Fetch and save Excel Header row text strings into @Headings array
    #
    @Headings = Spreadsheet::Read::row($bookdata->[1], 1);
    #
    #  Locate the Columns we will need to know by their Column Header Text
    #
    for $j (0 .. $MaxCols-1)
    {
        if ($Headings[$j] eq $VolFirstText)
        {
            $VolFirstCol = $j;                          # Save Column index for Volunteer First Name
        }
        if ($Headings[$j] eq $VolLastText)
        {
            $VolLastCol = $j;                           # Save Column index for Volunteer Last Name
        }
        if ($Headings[$j] eq $PhoneText)
        {
            $PhoneCol = $j;                             # Save Column index for Phone Number Called
        }
        if ($Headings[$j] eq $ResponseText)
        {
            $ResponseCol = $j;                          # Save Column index for Call Response
        }
        if ($Headings[$j] eq $DateText)
        {
            $DateCol = $j;                              # Save Column index for Call Date
        }
    }
    #
    #  Be Sure We Found the columns Headings We Need
    #
    if ($VolFirstCol == -1)
    {
        printLine("Couldn't Find Required Column Headed $VolFirstText -- Aborting...\n");
        goto EXIT;
    }
    if ($VolLastCol == -1)
    {
        printLine("Couldn't Find Required Column Headed $VolLastText -- Aborting...\n");
        goto EXIT;
    }
    if ($PhoneCol == -1)
    {
        printLine("Couldn't Find Required Column Headed $PhoneText -- Aborting...\n");
        goto EXIT;
    }
    if ($ResponseCol == -1)
    {
        printLine("Couldn't Find Required Column Headed $ResponseText -- Aborting...\n");
        goto EXIT;
    }
    if ($DateCol == -1)
    {
        printLine("Couldn't Find Required Column Headed $DateText -- Aborting...\n");
        goto EXIT;
    }
    printLine("... Located All Required Data Columns...\n");
    return;
}

#--------------------------------------------------------------------
# see if this file in the -survey directory qualifies to be selected
#
#  Filename has already been broken out into:
#       @phase = phase of this survey
#       @dist = AD or SD district of this survey
#       @party = party of the respondents to this survey
#       @likely = the voting propensity of the respondents
#
#  @opts = all -select options secified, one per array entry
#  @selphase = all phase -select criteria
#  @seldist = all district -select criteria
#  @selparty = all party -select criteria
#  @sellikely = all voting propensity -select criteria
#  
#--------------------------------------------------------------------
sub qualify_file {
    if (scalar @opts == 0) {
        return 0;                                                   # no sub-selection criteria, it qualifies
    }
    my $t = -1;                                                     # assume failure to select  
    my $x = 0;
    if (scalar @selphase > 0) {
        for $x (0 .. @selphase-1){                                  # phase parameter specified
            if ($selphase[$x] eq $phase) {
                $t=0;                                               # file Phase is selected
                last;                                               # move to any otehr params
            }
        }
        if ($t == -1) {
            return -1;                                              # file not selected
        }
    }
    $t = -1;
    if (scalar @seldist > 0) {                                      # district parameter specified
        for $x (0 .. @seldist-1){
            if ($seldist[$x] eq $dist) {
                $t=0;                                               # district is selected
                last;                                               # move to any otehr params
            }
        }
        if ($t == -1) {
            return -1;                                              # file not selected
        }
    }
    $t = -1;
    if (scalar @selparty > 0) {                                     # party parameter specified
        for $x (0 .. @selparty-1){
            if ($selparty[$x] eq $party) {
                $t=0;                                               # party is selected
                last;                                               # move to any otehr params
            }
        }
        if ($t == -1) {
            return -1;                                              # file not selected
        }
    }
    $t = -1;
    if (scalar @sellikely > 0) {                                    # propensity parameter specified
        for $x (0 .. @sellikely-1){
            if ($sellikely[$x] eq $likely) {
                $t=0;                                               # propensity is selected
                last;                                               # move to any otehr params
            }
        }
        if ($t == -1) {
            return -1;                                              # file not selected
        }
    }
    return 0;                                                       # file is selected
}

#--------------------------------------------------------------
# Print Logging line to both console and Log File
#--------------------------------------------------------------
sub printLine {
    my($printData) = @_;
    my $datestring = localtime();
    if ( substr( $printData , -1 ) ne "\r") {
        print $printFileh( PROGNAME . $datestring . ' ' . $printData);
    }
    print( PROGNAME . $datestring . ' ' . $printData );
}
