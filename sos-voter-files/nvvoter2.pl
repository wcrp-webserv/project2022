#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
# wcrp-voter-nvvoter2
#  merge voter rolls with votestat, emails, etc
#  Produce County base.csv file and precinct.csv file
#
#
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
use strict;
use warnings;
$| = 1;                                 # force STDOUT to flush on newline
use File::Basename;
use DBI;
use Data::Dumper;
use Getopt::Long qw(GetOptions);
use Time::Piece;
use Math::Round;
use Spreadsheet::Read;
use constant PROGNAME => "NVVOTER2 - ";
use Text::CSV qw( csv );

no warnings "uninitialized";

=head1 Function
=over
=head2 Overview
	This program will analyze a washoe-county-voter file
		a) file is sorted by precinct ascending
		b)
	Input: county voter registration file.
	       
	Output: a csv file containing the extracted fields 
=cut

my $records;

#
#  Configuration SpreadSheet File Name, Header & Data Row Arrays
#
my $CfgFile = "nvconfig.xlsx";                  # program configuration spreadsheet
my @CfgHeadings =();                            # Array of Text Headings for spreadsheet
my @CfgRow =();                                 # Data from the Row of spreadsheet currently being processed

# primary input from sec state
my $inputFile = "VoterList.ElgbVtr.43842.060420175555.csv";
my $inputFileh;
my $baseFile = "base.csv";
my $baseFileh;
my %baseLine = ();
my @baseLine;
my $baseLine;
my @baseProfile;

my $debugFile = "debug.txt";
my $debugFileh;
my $debug = 0;
my @duplicates;

# list of email addresses to add
my $voterEmailFile = "";
my $voterEmailFileh;
my @voterEmailArray;
my $voterEmailArray;
my @voterEmailHeadings;
my $voterEmailHeadings;

# email merge error report
my $emailLogFile = "email-adds-log.csv";
my $emailLogFileh;
my %emailLine = ();

# sorted voter statistic records (by voterid)
my $voterStatsFile = "voterdata-s.csv";
my $voterStatsFileh;
my %voterStatsArray;
my @voterStatsArray;
my $voterStatsArray;
my $voterStatsHeadings = "";
my @voterStatsHeadings;

my $printFile = "print.txt";
my $printFileh;

my $adPoliticalFile = "pctxref.csv";
my %politicalLine   = ();
my @adPoliticalHash = ();
my %adPoliticalHash;
my $adPoliticalHash;
my $adPoliticalHeadings = "";
my $precinctPolitical;
my @precinctPolitical;
my $noPoliticalWarn = 0;
my $Noxref = 0;

my $stats;
my $emails;

my @electionValue = ();
my $regTimePiece = 0;

my $helpReq     = 0;
my $maxLines    = "250";
my $voteCycle   = "";
my $fileCount   = 1;
my $csvHeadings = "";
my @csvHeadings;
my $line1Read    = '';
my @line1Read;
my $linesRead    = 0;
my $linesIncRead = 0;
my $printData;
my $linesWritten = 0;
my $emailAdded   = 0;
my $statsAdded   = 0;
 my $voterid = '';

my $skipRecords    = 20;
my $skippedRecords = 0;

my $generalCount;
my $party;
my $primaryCount;
my $pollCount;
my $absenteeCount   = 0;
my $activeVOTERS    = 0;
my $activeREP       = 0;
my $activeDEM       = 0;
my $activeOTHR      = 0;
my $totalVOTERS     = 0;
my $totalAMER       = 0;
my $totalAMEL       = 0;
my $totalDEM        = 0;
my $totalDUO        = 0;
my $totalFED        = 0;
my $totalGRN        = 0;
my $totalIA         = 0;
my $totalIAP        = 0;
my $totalIND        = 0;
my $totalINAC       = 0;
my $totalLIB        = 0;
my $totalLPN        = 0;
my $totalNL         = 0;
my $totalNP         = 0;
my $totalORGL       = 0;
my $totalOTH        = 0;
my $totalPF         = 0;
my $totalPOP        = 0;
my $totalREP        = 0;
my $totalRFM        = 0;
my $totalSOC        = 0;
my $totalTEANV      = 0;
my $totalUWS        = 0;
my $totalGENERALS   = 0;
my $totalPRIMARIES  = 0;
my $totalPOLLS      = 0;
my $totalABSENTEE   = 0;
my $totalSTRDEM     = 0;
my $totalMODDEM     = 0;
my $totalWEAKDEM    = 0;
my $percentSTRGRDEM = 0;
my $totalSTRREP     = 0;
my $totalMODREP     = 0;
my $totalWEAKREP    = 0;
my $percentSTRGREP  = 0;
my $totalSTROTHR    = 0;
my $totalMODOTHR    = 0;
my $totalWEAKOTHR   = 0;
my $percentSTRGOTHR = 0;
my $totalOTHR       = 0;
#
# Precinct file data
#
#  $NumPct = Number of precincts (and thus the number of entries in each parallel precinct array)
#  @PctPrecinct = Array of Precinct Numbers in this base.csv compilation
#  @Pctxxx = parallel Array of counts for each item by precinct
#
my $pctFile = "precinct.csv";
my $pctFileh;
my $NumPct = 0;                     # Number of Precincts
my @PctPrecinct =();                # Array of precinct numbers
my @PctCD = ();                     # Array of Congressional Districts
my @PctAD =();                      # Array of State Assembly Districts
my @PctSD =();                      # Array of State Senate Districts
my @PctBoardofEd =();                 # Array of Board of Ed
my @PctCntyComm =();                 # Array of Board of Ed
my @PctRwards =();                 # Array of Board of Ed
my @PctSwards =();                 # Array of Board of Ed
my @PctSchBdTrust =();                 # Array of Board of Ed
my @PctSchBdAtLrg =();                 # Array of Board of Ed
my @PctGenerals =();                # Total general election votes this precinct
my @PctPrimaries =();               # Total primary election votes this precinct
my @PctPolls =();                   # Total poll votes this precinct
my @PctAbsentee =();                # Total mail in votes this precinct
my @PctRegRep =();                  # Array of # Registered Republicans
my @PctRegDem =();                  # Array of # Registered Democrats
my @PctRegNP =();                   # Array of # Registered Non-Partisans
my @PctRegIAP =();                  # Array of # Registered Independent American Party
my @PctRegLP =();                   # Array of # Registered Libertarian Party
my @PctRegGP =();                   # Array of # Registered Green Party
my @PctRegOther =();                # Array of # Registered to Other Parties
my @PctStrongRep =();               # Array of # Strong Voting Republicans
my @PctModRep =();                  # Array of # Moderate Voting Republicans
my @PctWeakRep =();                 # Array of # Weak Voting Republicans
my @PctStrongDem =();               # Array of # Strong Voting Democrats
my @PctModDem =();                  # Array of # Moderate Voting Democrats
my @PctWeakDem =();                 # Array of # Weak Voting Democrats
my @PctStrongAllOther =();          # Array of # Strong Voting All Other Parties
my @PctModAllOther =();             # Array of # Moderate Voting All Other Parties
my @PctWeakAllOther =();            # Array of # Weak Voting All Other Parties
my @PctActiveRep =();               # Array of # of active Republican
my @PctActiveDem =();               # Array of # of active Democrat
my @PctActiveAllOther =();          # Array of # of active voter in All Other Parties
#
my $pctHeading = "";
my @pctHeading =(                   # Header for precinct.csv
                "Precinct",         # Precinct Number
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
                "#Weak Other"       # Total Weak All Other Parties
);

my @csvRowHash;
my %csvRowHash = ();
my @partyHash;
my %partyHash  = ();
my %schRowHash = ();
my @schRowHash;
my @values1;
my @values2;
my $voterRank;

my $calastName;
my $cafirstName;
my $camiddleName;
my $caemail;
my $capoints;

my $baseHeading = "";
my $fixedflds = 32;                         # 32 fixed fields before votedata
my @baseHeading = (                 # base.csv file header
    "CountyID",     "StateID",  "Status",   "County",    "Precinct", "CongDist",
    "AssmDist",     "SenDist",  "BrdofEd",  "CntyComm",  "Rwards",   "Swards",   "SchBdTrust", "SchBdAtLrg",
    "First",        "Last",     "Middle",   "Suffix",    "Phone",    "email",
    "BirthDate",    "RegDate",  "Party",    "StreetNo",
    "StreetName",   "Address1", "Address2", "City",
    "State",        "Zip",   
    "RegisteredDays", "Age", 
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
    "LikelytoVote", "Score",
);
my @emailProfile;
my $emailHeading = "";
my @emailHeading =
  ( "VoterID", "Precinct", "First", "Last", "Middle", "email", );

my @votingLine;
my $votingLine;
my @votingProfile;

my $precinct = "000000";
my $noVotes  = 0;
my $noData   = 0;

#
#  Array that will be loaded with the highest voter ID that voted in each of the 20 elections being tracked
#
my @HighVoterIDs = (
    -1,                                         # -1 indicates not yet loaded
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0     # Highest Voter ID that voted in the 20 elections (inits at 0)
);

#
# main program controller
#
sub main {

    #Open file for messages and errors
    open( $printFileh, '>>', "$printFile" )
      or die "Unable to open PRINT: $printFile Reason: $!";

    # Parse any parameters
    GetOptions(
        'infile=s'    => \$inputFile,
        'outfile=s'   => \$baseFile,
        'pctfile=s'   => \$pctFile,
        'config=s'    => \$CfgFile,
        'statfile=s'  => \$voterStatsFile,
        'emailfile=s' => \$voterEmailFile,
        'xref=s'      => \$adPoliticalFile,
        'help!'       => \$helpReq,
        'h'           => \$helpReq
    ) or $helpReq = 1;

    my $csv = Text::CSV->new(
        {
            binary             => 1,  # Allow special character. Always set this
            auto_diag          => 1,  # Report irregularities immediately
            allow_whitespace   => 0,
            allow_loose_quotes => 1,
            quote_space        => 0,
        }
    );

    if ($helpReq) {
        print ("perl nvvoter2 -infile <file> -outfile <file> -pctfile<file> -statfile <file> -xref <file> -emailfile <file>\n");
        print ("    -infile = Secretary Of State ElgbVtr file\n");
        print ("    -outfile = compiled \"base\" file - default is base.csv\n");
        print ("    -pctfile = precinct summary file - default is precinct.csv\n");
        print ("    -statfile = input by voter vote records from NVVOTER1 - default is voterdata-s.csv\n");
        print ("    -xref = Precinct to political district cross reference file - default is pctxref.csv\n");
        print ("    -emailfile = optional file of email addresses to add to base.csv on name match\n");
        die "\n";
    }

    load_config();                       # load configuration spreadsheet

    printLine("My inputfile is: $inputFile.\n");
    open( $inputFileh, $inputFile )
        or die ("Unable to open INPUT: $inputFile Reason: $!\n");

    # Read in the header line
    $line1Read = $csv->getline($inputFileh);

    # move headings into an array to modify
    @csvHeadings = @$line1Read;

    # Remove spaces in heading text
    my $j = @csvHeadings;
    for ( my $i = 0 ; $i < ( $j - 1 ) ; $i++ ) {
        $csvHeadings[$i] =~ s/\s//g;
        #
        #  If heading entry begins with "Residential " remove that part
        #
        $csvHeadings[$i] =~ s/Residential//;
    }
    # $csvHeadings is now the header of the incoming voterdata-s.csv file ready to hash with each data line

    # Build header for optional email file
    $emailHeading = join( ",", @emailHeading );
    $emailHeading = $emailHeading . "\n";

    # Open base.csv output file
    printLine("Voter Base-table file: $baseFile\n");
    open( $baseFileh, ">$baseFile" )
      or die "Unable to open baseFile: $baseFile Reason: $!";
   
    # Build header for base.csv file record and write it out
    $baseHeading = join( ",", @baseHeading );
    $baseHeading = $baseHeading . "\n";
    print $baseFileh $baseHeading;

    # Open precinct.csv file
    printLine("Voter precinct-table file: $pctFile\n");
    open( $pctFileh, ">$pctFile" )
      or die "Unable to open baseFile: $pctFile Reason: $!";
    
    # Build and write out header for precinct.csv file
    $pctHeading = join( ",", @pctHeading );
    $pctHeading = $pctHeading . "\n";
    print $pctFileh $pctHeading;

    # initialize the precinct-all table
    printLine ("Build Political District Hash \n");
    adPoliticalAll(@adPoliticalHash);

    # initialize the optional voter email log and the email array if selected
    if ( $voterEmailFile ne "" ) {
        printLine("My emailFile is: $voterEmailFile.\n");
        printLine("Email updates file: $voterEmailFile\n");
        open( $emailLogFileh, ">$emailLogFile" )
          or die "Unable to open emailLogFileh: $emailLogFile Reason: $!";
        print $emailLogFileh $emailHeading;
        voterEmailLoad(@voterEmailArray);
    }

    # if voter stats are available load the hash table
    if ( $voterStatsFile ne "" ) {
        printLine("Vote History file: $voterStatsFile\n");
        voterStatsLoad(@voterStatsArray);
    }

    #----------------------------------------------------------
    # Process loop
    # Read the entire input and
    # 1) edit the input lines
    # 2) transform the data
    # 3) write out transformed line
    #----------------------------------------------------------

  NEW:
    my $tmp = "";
    while ( $line1Read = $csv->getline($inputFileh) ) {
        $linesRead++;
        $linesIncRead++;
        if ( $linesIncRead > 4999 ) {
            printLine("$linesRead voter records read\r");
            $linesIncRead = 0;
        }

        # create the values array to complete preprocessing
        @values1 = @$line1Read;
        @csvRowHash{@csvHeadings} = @values1;

        #- - - - - - - - - - - - - - - - - - - - - - - - - -
        # Assemble database load  for base segment
        #- - - - - - - - - - - - - - - - - - - - - - - - - -
        %baseLine = ();

        $baseLine{"StateID"} = $csvRowHash{"VoterID"};
        $voterid = $csvRowHash{"VoterID"};

        $baseLine{"CountyID"} = $csvRowHash{"CountyVoterID"};
        $baseLine{"Status"}   = $csvRowHash{"CountyStatus"};
        $baseLine{"County"} = $csvRowHash{"County"};
        $baseLine{"Precinct"} = $csvRowHash{"RegisteredPrecinct"};
        $baseLine{"CongDist"} = $csvRowHash{"CongressionalDistrict"};
        $baseLine{'AssmDist'} = $csvRowHash{"AssemblyDistrict"};
        $baseLine{'SenDist'}  = $csvRowHash{"SenateDistrict"};
       
        my $precinct = $csvRowHash{"RegisteredPrecinct"};           # get this voter's precinct
        if ($Noxref == 0) {
            $precinctPolitical = $adPoliticalHash{$precinct};           # fetch XREF array reference for this precinct
            my $test = $precinctPolitical->[0];                         # see if an XREF entry was found
            if (!defined $test || $test eq "") {
                if ($noPoliticalWarn == 0) {
                    #
                    #  No XREF entry for this precinct
                    #
                    printLine ("******** WARNING!! YOU NEED TO UPDATE PRECINCT XREF FILE \n");
                    printLine ("******** At least Precinct $precinct not in precinct xref file\n");
                    printLine ("******** File debug.txt lists all missing precincts.\n");
                    #
                    #  Open debug.txt to list missing precincts
                    #
                    $debug = 1;
                    open($debugFileh, ">", $debugFile )
                    or $debug = 0;                                  # disable if for some reason doesn't open
                    if ($debug == 0) {
                        printLine (">>>>>>>> Could Not Create debug.txt file\n");
                    }
                    $noPoliticalWarn = 1;                           # don't warn again of missing precincts on console
                }
                if ($debug != 0) {
                    #
                    #  List all missing precincts in debug.txt, but not duplicates (same precinct missing in more than one voter record)
                    #
                    my $dup = 0;
                    for ( my $i = 0 ; $i <= $#duplicates ; $i++ ) {
                        if ( $precinct == $duplicates[$i]) {
                            $dup = 1;                                       # already listed, skip listing it again
                            last;
                        }
                    }
                    if ($dup == 0) {
                        #
                        #  List and remember a new missing precinct in debug.txt
                        #
                        push @duplicates, $precinct;                        # add to duplicate missing precinct detection list
                        print $debugFileh ("Precinct $precinct not in precinct xref file\n");
                    }
                }
            } else {
                #
                #  Found an XREF record for this precinct, Fill in the political districts from the XREF file
                #
                $baseLine{"SenDist"}       = "SD" . $precinctPolitical->[1];
                $baseLine{"AssmDist"}      = "AD" . $precinctPolitical->[2];
                $baseLine{"BrdofEd"}       = $precinctPolitical->[3];
                $baseLine{"CntyComm"}      = $precinctPolitical->[5];
                $baseLine{"Rwards"}        = $precinctPolitical->[6];
                $baseLine{"Swards"}        = $precinctPolitical->[7];
                $baseLine{"SchBdTrust"}    = $precinctPolitical->[8];
                $baseLine{"SchBdAtLrg"}    = $precinctPolitical->[9];
            }
        }
        # convert proper names to upper case first then lower
        my $UCword = $csvRowHash{"FirstName"};
        $UCword =~ s/(\w+)/\u\L$1/g;
        $baseLine{"First"} = $UCword;
        my $ccfirstName = $UCword;

        $UCword = $csvRowHash{"MiddleName"};
        $UCword =~ s/(\w+)/\u\L$1/g;
        $baseLine{"Middle"} = $UCword;
        $UCword = $csvRowHash{"LastName"};

        $UCword =~ s/(\w+)/\u\L$1/g;
        if ( $UCword =~ m/,/ ) {
            $UCword =~ s/\s+//g;     # remove all imbedded spaces
            $UCword =~ s/,/-/g;      # change comma to dash
        }
        $baseLine{"Last"} = $UCword;
        my $cclastName = $UCword;
        $UCword =~ s/(\w+)/\u\L$1/g;

        $baseLine{"BirthDate"} = $csvRowHash{"BirthDate"};
        $baseLine{"RegDate"}   = $csvRowHash{"RegistrationDate"};
        $baseLine{"Party"}     = $csvRowHash{"Party"};
        $baseLine{"Phone"}     = $csvRowHash{"Phone"};
        $UCword                = $csvRowHash{"Address1"};
        $UCword =~ s/(\w+)/\u\L$1/g;
        $baseLine{"Address1"} = $UCword;
        my @streetno = split( / /, $UCword, 2 );
        $baseLine{"StreetNo"}   = $streetno[0];
        $baseLine{"StreetName"} = $streetno[1];
        $UCword                 = $csvRowHash{"City"};
        $UCword =~ s/(\w+)/\u\L$1/g;
        $baseLine{"City"}  = $UCword;
        $baseLine{"State"} = $csvRowHash{"State"};
        $baseLine{"Zip"}   = $csvRowHash{"Zip"};
        $baseLine{"email"} = "";
       
        #
        #  locate and add voter statistics
        $stats = -1;
        $stats = binary_search( \@voterStatsArray, $voterid );
        if ( $stats != -1 ) {
            for ( my $i = 1 ; $i <= 29 ; $i++ ) {
                # copy over 29 fields from voterdata.csv file record
                $baseLine{ $baseHeading[ $i + $fixedflds-1 ] } = $voterStatsArray[$stats][$i];  ###
            }
            $statsAdded++;
        }
        else {
            # fill in record for registered voter with no vote history
            $noData++;
            for ( my $i = 1 ; $i <= 20 ; $i++ ) {
                $baseLine[ $i + $fixedflds ] = "";    # blank all 20 election votes
            }
            $baseLine{"Generals"}     = 0;
            $baseLine{"Primaries"}    = 0;
            $baseLine{"Polls"}        = 0;
            $baseLine{"Absentee"}     = 0;
            $baseLine{"Early"}        = 0;
            $baseLine{"Provisional"}  = 0;
            $baseLine{"LikelytoVote"} = "WEAK";
            $baseLine{"Score"}        = 0;
            $baseLine{"TotalVotes"}   = 0;
        }
        if ( $baseLine{"TotalVotes"} == 0 ) {
            $noVotes++;
        }

        # calc age and registration days
        calc_days();
#
#  locate email address if available
#  "Last", "First", "Middle","Phone","email","Address", "City","Contact Points",
#     0       1         2                4      5          6          7
        $emails = binary_ch_search( \@voterEmailArray, $cclastName );
        if ( $emails != -1 ) {
            printLine("Email index = $emails not -1\n");
            if (   $voterEmailArray[$emails][0] eq $cclastName
                && $voterEmailArray[$emails][1] eq $ccfirstName )
            {
                $calastName        = $voterEmailArray[$emails][0];
                $cafirstName       = $voterEmailArray[$emails][1];
                $caemail           = $voterEmailArray[$emails][4];
                $baseLine{"email"} = $voterEmailArray[$emails][4];
                $capoints          = $voterEmailArray[$emails][7];
                $capoints =~ s/;/,/g;
                $emailAdded = $emailAdded + 1;

                # build a trace line to show email was updated
                %emailLine = ();
                $emailLine{"VoterID"} = $voterid;
                $emailLine{"Precinct"} = substr $csvRowHash{"RegisteredPrecinct"}, 0, 6;
                $emailLine{"Last"}     = $calastName;
                $emailLine{"First"}    = $cafirstName;
                $emailLine{"email"}    = $caemail;
                @emailProfile          = ();

                foreach (@emailHeading) {
                    push( @emailProfile, $emailLine{$_} );
                }
                print $emailLogFileh join( ',', @emailProfile ), "\n";
            }
        }
        #
        # Finally - output finished base.csv row from $baseLine
        #
        @baseProfile = ();
        foreach (@baseHeading) {
            if ($baseLine{$_} =~ /[\"\',]/) {
                $baseLine{$_} = "\"".$baseLine{$_}."\"";            # Quote  entry if contains comma or quote mark
            }
            push( @baseProfile, $baseLine{$_} );                    # build output line in baseProfile
        }
        print $baseFileh join( ',', @baseProfile ), "\n";           # write this row to base.csv file
        $linesWritten++;
        #
        #  Do the precinct stats accumulation for this record
        #
        calc_precinct();
        #
        # For now this is the in-elegant way I detect completion
        #
        if ( eof($inputFileh) ) {
            write_precinct();                   # write the precinct.csv file
            goto EXIT;
        }
        next;
    }
    #
    goto NEW;
}
#
# call main program controller
main();
#
# Common Exit
EXIT:

printLine("<===> Completed transformation of: $inputFile \n");
printLine("<===> Total Eligible Voter Records Read: $linesRead \n");
printLine("<===> Total Voting History Stats added: $statsAdded \n");
printLine("<===> Total Registered Voters with no Recent Vote History: $noVotes\n");
printLine("<===> Total Registered Voters with no Vote Record: $noData\n");
printLine("<===> Total Email Addresses added: $emailAdded \n");
printLine("<===> Total Precincts found and Precinct.csv Records written: $NumPct \n");
printLine("<===> Total base.csv Records written: $linesWritten \n");

close($inputFileh);
close($baseFileh);
close($pctFileh);
close($printFileh);
if ($debugFileh != 0) {
    close($debugFileh);
}
if ( $voterEmailFile ne "" ) {
    close($emailLogFileh);
}
exit;

#---------------------------------------------------------------
#
#  Routine to calculate age & registered days
#
sub calc_days {
    my $birthdate = $csvRowHash{"BirthDate"};
    my $regdate   = $csvRowHash{"RegistrationDate"};
    my $adjDate;

    # determine age
    my ( @date, $yy, $mm, $dd, $now, $age, $regdays, $before, $adjustedDate );
    if ( $birthdate ne "" ) {
        @date         = split( /\s*\/\s*/, $birthdate, -1 );
        $mm           = sprintf( "%02d", $date[0] );
        $dd           = sprintf( "%02d", $date[1] );
        $yy           = sprintf( "%02d", substr( $date[2], 0, 4 ) );
        $adjustedDate = "$mm/$dd/$yy";                          
        my $datelen =length($adjustedDate);
        if ($datelen <= 8) {
            $before    = Time::Piece->strptime( $adjustedDate, "%m/%d/%y" );  
        } else {
            $before    = Time::Piece->strptime( $adjustedDate, "%m/%d/%Y" );  
        }        
#        printLine("line 481 $adjDate: $adjdDate  \n");
#        $before       = Time::Piece->strptime( $adjDate, "%m/%d/%Y" );
        $now          = localtime;
        $age          = $now - $before;
        $age          = ( $age / (86400) / 365 );
        $age          = round($age);
    }
    else {
        $age = "";
    }
    $baseLine{"Age"} = $age;
    # determine registered days
    # may get dates in two formats: mm/dd/yyyy or yyyy-mm-dd
    if ( substr( $regdate, 4, 1 ) eq '-' ) {

        # handle yyyy-mm-dd (ISO-8898)
        @date = split( /\s*\-\s*/, $regdate, -1 );
        $mm   = $date[1];
        $dd   = $date[2];
        $yy   = $date[0];
    }
    else {
        # handle mm/dd/yyyy
        @date = split( /\s*\/\s*/, $regdate, -1 );
        $mm   = sprintf( "%02d", $date[0] );
        $dd   = sprintf( "%02d", $date[1] );
        $yy   = sprintf( "%02d", substr( $date[2], 0, 4 ) );
    }

    if ( $yy < 1900 ) {
        $yy = 2016;
    }
    $adjustedDate = "$mm/$dd/$yy";
#    printLine("line 648 adjustedDate:  $adjustedDate  \n");

    $before  = Time::Piece->strptime( $adjustedDate, "%m/%d/%Y" );
    $regTimePiece = $before;                                            # save encoded registration date for later work
    $now     = localtime;
    $regdays = $now - $before;
    $regdays = ( $regdays / (86400) );
    $regdays = round($regdays);
    $baseLine{"RegisteredDays"} = $regdays;
    #
    #  Find oldest election reg date allows vote in.  If older vote, use that date
    #  as it means voter re-registered at some point.
    #
    my $rstop = 0;
    my $ovote="";
    my $test = 0;
    my $vid = $csvRowHash{"VoterID"};
    for my $j (0 .. 19) {
        if ($rstop == 0) {
            my $edate = Time::Piece->strptime( substr($baseHeading[$j+$fixedflds], 0 , 8), "%m/%d/%y" );   # election date
            if ($edate < $regTimePiece) {
                $rstop = $j;                                                            # index+1 to oldest election registered for
            }
        }else{
            #
            #  See if older vote than registration date
            #
            if ($baseLine{$baseHeading[$j+$fixedflds]} ne "") {
                $rstop = $j;                                                             # must have re-registered
                $test = 1;
            }
        }
    }
    #
    #  $rstop = index to oldest possible vote for this voter.
    #  calculate voter propensity to vote strength based on
    #  this many possible votes.
    #
    my $maxstrength = 0;
    my $voterstrength = 0;                      # init accumulators
    for my $j (0 .. $rstop) {
        $maxstrength = $maxstrength + $electionValue[$j+1];               # sum possible election strengths
        if ($baseLine{$baseHeading[$j+$fixedflds]} ne "") {
            $voterstrength = $voterstrength + $electionValue[$j+1];       # sum actual voted election strengths
        }
    }
    $voterstrength = (($voterstrength/$maxstrength) * 10);                # calc voter strength 0-9.99
    $baseLine{"Score"} = $voterstrength;
    if ($voterstrength <= 2) {
        $baseLine{"LikelytoVote"} = "WEAK";                                             # < 2 = weak
    }
    if ($voterstrength == 0) {
        $baseLine{"LikelytoVote"} = "NEVER";                                            # special case strength of 0
    }
    if ($baseLine{"Primaries"} > 0) {
        $baseLine{"LikelytoVote"} = "MODERATE";                                         # Moderate if voted in a primary
    }
    if (($voterstrength > 2) and ($voterstrength <= 6)) {
        $baseLine{"LikelytoVote"} = "MODERATE";
    }
    if ($voterstrength > 6) {
        $baseLine{"LikelytoVote"} = "STRONG";
    }
    $voterstrength = int($voterstrength + 0.49);                                         # convert to 0-10 score
    $baseLine{"Score"} = $voterstrength;
}

#---------------------------------------------------------------
#   Numeric Binary Search
#
# $index = binary_search( \@array, $word )
#   @array is a list of lowercase strings in alphabetical order.
#   $word is the target word that might be in the list.
#   binary_search() returns the array index such that $array[$index]
#   is $word.
sub binary_search {
    my ( $try,   $var );
    my ( $array, $word ) = @_;
    my ( $low,   $high ) = ( 0, @$array - 1 );
    while ( $low <= $high ) {    # While the window is open
        $try = int( ( $low + $high ) / 2 );    # Try the middle element
        $var = $array->[$try][0];
        $low  = $try + 1, next if $array->[$try][0] < $word;    # Raise bottom
        $high = $try - 1, next if $array->[$try][0] > $word;    # Lower top
        return $try;    # We've found the word!
    }
    $try = -1;
    return $try;        # The word isn't there.
}

#---------------------------------------------------------------
#
# binary search for character strings
#
sub binary_ch_search {
    my ( $try,   $var );
    my ( $array, $word ) = @_;
    my ( $low,   $high ) = ( 0, @$array - 1 );
    while ( $low <= $high ) {    # While the window is open
        $try = int( ( $low + $high ) / 2 );    # Try the middle element
        $var = $array->[$try][0];
        $low  = $try + 1, next if $array->[$try][0] lt $word;    # Raise bottom
        $high = $try - 1, next if $array->[$try][0] gt $word;    # Lower top
        return $try;    # We've found the word!
    }
    $try = -1;
    return $try;        # The word isn't there.
}

#---------------------------------------------------------------
#
#  Calculate data for precinct.csv file
#  Called for each line in S.O.S. data file after
#  all processing to create $baseLine Hash for this record
#
sub calc_precinct() {
    my $i =0;
    my $Active=0;
    if ( $NumPct == 0) {
        add_pct();
        $i=0                                        #1st precinct added, set index
    } else {  
        for ( $i = 0 ; $i < ( $NumPct) ; $i++ ) {
            if ($PctPrecinct[$i] == $baseLine{"Precinct"}) {
                last;
            }
        }
        if ($i == $NumPct) {
            add_pct();                              # new precinct, add a row for it
        }
    }
    #
    #  $i now = index for this precinct's row in the precinct parallel array matrix
    #
    #  Accumulate the stats from this voter's $baseLine data.
    #
    $PctGenerals[$i] = $PctGenerals[$i] + $baseLine{"Generals"};
    $PctPrimaries[$i] = $PctPrimaries[$i] + $baseLine{"Primaries"};
    $PctPolls[$i] = $PctPolls[$i] + $baseLine{"Polls"};
    $PctAbsentee[$i] = $PctAbsentee[$i] + $baseLine{"Absentee"};
    $Active=0;                                      # Assume Inactive Voter
    if ($baseLine{"Status"} eq "Active") {
        $Active = 1;                                # set 1 more Active Voter
    }
    if ($baseLine{"Party"} eq "Republican") {
        #
        #  process Republican Voter
        #
        $PctRegRep[$i]++;                           # Count another Registered Republican
        $PctActiveRep[$i] = $PctActiveRep[$i] + $Active; # accumulate # active republican voters in precinct
        if ($baseLine{"LikelytoVote"} eq "STRONG") {
            $PctStrongRep[$i]++;                    # Count as strong republican
        }
        if ($baseLine{"LikelytoVote"} eq "MODERATE") {
            $PctModRep[$i]++;                       # Count as moderate republican
        }
        if ($baseLine{"LikelytoVote"} eq "WEAK") {
            $PctWeakRep[$i]++;                      # Count as weak republican
        }
        return;                                     # done with this voter
    }
    if ($baseLine{"Party"} eq "Democrat") {
        #
        #  process DEmocrat Voter
        #
        $PctRegDem[$i]++;                           # Count another Registered Democrat
        $PctActiveDem[$i] = $PctActiveDem[$i] + $Active; # accumulate # active Democrat voters in precinct
        if ($baseLine{"LikelytoVote"} eq "STRONG") {
            $PctStrongDem[$i]++;                    # Count as strong Democrat
        }
        if ($baseLine{"LikelytoVote"} eq "MODERATE") {
            $PctModDem[$i]++;                       # Count as moderate Democrat
        }
        if ($baseLine{"LikelytoVote"} eq "WEAK") {
            $PctWeakDem[$i]++;                      # Count as weak Democrat
        }
        return;                                     # done with this voter
    }
    #
    #  Voter is not Republican or Democrat, so do the All OTHER PARTY stats\
    #
    $PctActiveAllOther[$i] = $PctActiveAllOther[$i] + $Active; # accumulate # active All non dem or Rep Party voters in precinct
    if ($baseLine{"LikelytoVote"} eq "STRONG") {
        $PctStrongAllOther[$i]++;                   # Count as strong Other
        }
        if ($baseLine{"LikelytoVote"} eq "MODERATE") {
            $PctModAllOther[$i]++;                  # Count as moderate Other
        }
        if ($baseLine{"LikelytoVote"} eq "WEAK") {
            $PctWeakAllOther[$i]++;                 # Count as weak Other
        }
    #
    #  Now Try to Find which OTHER party we might care about
    #
    if ($baseLine{"Party"} eq "Independent American Party") {
        $PctRegIAP[$i]++;
        return;
    }
    if ($baseLine{"Party"} eq "Green Party") {
        $PctRegGP[$i]++;
        return;
    }
    if ($baseLine{"Party"} eq "Non-Partisan") {
        $PctRegNP[$i]++;
        return;
    }
    if ($baseLine{"Party"} eq "Libertarian Party") {
        $PctRegLP[$i]++;
        return;
    } 
    $PctRegOther[$i]++;                            # Count as Registered some Other Party
    return;
}

#---------------------------------------------------------------
#
#  Add a new precinct row to the parallel precinct tables
#
sub add_pct() {
    $NumPct = $NumPct+1;                               # add an array row
    push(@PctPrecinct, $baseLine{"Precinct"});       # set precinct number
    push(@PctCD,  $baseLine{"CongDist"});            # set CD for this precinct
    push(@PctAD, $baseLine{'AssmDist'});             # set Assembly District
    push(@PctSD, $baseLine{'SenDist'});              # set Senate District
    push(@PctBoardofEd, $baseLine{'BrdofEd'});         # set Board of Education
    push(@PctCntyComm, $baseLine{'CntyComm'});         # set Board of Education
    push(@PctRwards, $baseLine{'Rwards'});         # set Board of Education
    push(@PctSwards, $baseLine{'Swards'});         # set Board of Education
    push(@PctSchBdTrust, $baseLine{'SchBdTrust'});         # set Board of Education
    push(@PctSchBdAtLrg, $baseLine{'SchBdAtLrg'});         # set Board of Education
    push(@PctGenerals, 0);                           # init rest of row's data to zeroes
    push(@PctPrimaries, 0);
    push(@PctPolls, 0);
    push(@PctAbsentee, 0);
    push(@PctRegRep, 0);                             # init rest of row's data to zeroes
    push(@PctRegDem, 0);
    push(@PctRegNP, 0);
    push(@PctRegIAP, 0);
    push(@PctRegLP, 0);
    push(@PctRegGP, 0);
    push(@PctRegOther, 0);
    push(@PctStrongRep, 0);
    push(@PctModRep, 0);
    push(@PctWeakRep, 0);
    push(@PctStrongDem, 0);
    push(@PctModDem, 0);
    push(@PctWeakDem, 0);
    push(@PctStrongAllOther, 0);
    push(@PctModAllOther, 0);
    push(@PctWeakAllOther, 0);
    push(@PctActiveRep, 0);
    push(@PctActiveDem, 0);
    push(@PctActiveAllOther, 0);
}

#---------------------------------------------------------------
#
#  Write out the precinct matrix data to the precinct.csv file
#
sub write_precinct() {
    my $lineout;
    my $totvote;
    my $pctRep;
    my $pctDem;
    my $pctAllOther;
    my $numAllOther;
    my $j;
    my $i;
    my @PctSort = ();
    @PctSort = sort { $a <=> $b } @PctPrecinct;                             # get precionct numbers in ascending order
    for ( $j = 0 ; $j < $NumPct ; $j++ ) {
        for ( $i = 0 ; $i < $NumPct ; $i++) {
            if ($PctPrecinct[$i] == $PctSort[$j]) {
                last;                                                       # $i pionts to next ascending precinct
            }
        }
        # calc  voters registereed to all other parties in precinct
        my $numAllOther = $PctRegNP[$i] + $PctRegIAP[$i] + $PctRegLP[$i] + $PctRegGP[$i] + $PctRegOther[$i];
        $totvote = $PctRegRep[$i] + $PctRegDem[$i] + $numAllOther;          # Calc Total Voters in precinct
        if ($totvote == 0) {
            $totvote = 1;                                                   # avoid divide by zero if no voters in a precinct
        }
        $pctRep = int((($PctRegRep[$i] / $totvote) * 10000)+.5)/100;        # percent of precinct republican 
        $pctDem = int((($PctRegDem[$i] / $totvote) * 10000)+.5)/100;        # percent of precinct democrat
        $pctAllOther = int((($numAllOther / $totvote) * 10000)+.5)/100;     # percent of precinct All Other Party Registration
        #
        #  There's probably a better way to build the output line, but I don't know
        #  what it is so here goes brute force.
        #
        $lineout = $PctPrecinct[$i] . "," . $PctCD[$i] . "," . $PctAD[$i] . "," . $PctSD[$i] . ","; 
        $lineout = $lineout . $PctBoardofEd[$i] . "," . $PctCntyComm[$i] . "," . $PctRwards[$i] . ",";
        $lineout = $lineout . $PctSwards[$i] . "," . $PctSchBdTrust[$i] . "," . $PctSchBdAtLrg [$i] . ",";
        $lineout = $lineout . $PctGenerals[$i] . "," . $PctPrimaries[$i] . ",";
        $lineout = $lineout . $PctPolls[$i] . "," . $PctAbsentee[$i] . "," . $PctRegNP[$i] . ",";
        $lineout = $lineout . $PctRegIAP[$i] . "," . $PctRegLP[$i] . "," . $PctRegGP[$i] . ","  . $PctRegOther[$i] . ",";
        $lineout = $lineout . $PctRegRep[$i] . "," . $PctActiveRep[$i] . "," . $pctRep . "%,";
        $lineout = $lineout . $PctRegDem[$i] . "," . $PctActiveDem[$i] . "," . $pctDem . "%,";
        $lineout = $lineout . $numAllOther . "," . $PctActiveAllOther[$i] . "," . $pctAllOther . "%,";
        $lineout = $lineout . $PctStrongRep[$i] . "," . $PctModRep[$i] . "," . $PctWeakRep[$i] . ",";
        $lineout = $lineout . $PctStrongDem[$i] . "," . $PctModDem[$i] . "," . $PctWeakDem[$i] . ",";
        $lineout = $lineout . $PctStrongAllOther[$i] . "," . $PctModAllOther[$i] . "," . $PctWeakAllOther[$i] . ",";

        $lineout = $lineout  . "\n";
        print $pctFileh $lineout;
    } 
}

#---------------------------------------------------------------
#
# Load the Vote History array that will be accessed via binary search
#
sub voterStatsLoad() {
    printLine("Started building Vote History hash \n");

    my $loadCnt = 0;
    my $Scsv    = Text::CSV->new(
        {
            binary             => 1,  # Allow special character. Always set this
            auto_diag          => 1,  # Report irregularities immediately
            allow_whitespace   => 0,
            allow_loose_quotes => 1,
            quote_space        => 0,
        }
    );

    $voterStatsHeadings = "";
    open( $voterStatsFileh, $voterStatsFile )
      or die "Unable to open INPUT: $voterStatsFile Reason: $!";

    $line1Read = $Scsv->getline($voterStatsFileh);    # get header
    @voterStatsHeadings = @$line1Read;    # in voter Stats Headings Array

                                          # Build the UID->survey hash
    while ( $line1Read = $Scsv->getline($voterStatsFileh) ) {
        if ($line1Read->[0] == 0 ) {
            #
            # This is the Highest Voter ID in each election record
            # copy to @HighVoterID array
            #
            $HighVoterIDs[0] = 0;                       # indicate data loaded
            for my $z (1 .. 20) {
                $HighVoterIDs[$z] = $line1Read->[$z];   # load values into array
#               printLine("... $voterStatsHeadings[$z] - High ID = $HighVoterIDs[$z] \n")
            }
        } else {
            #
            # This is a normal voter vote data record, add to voterStatsArray
            #
            my @values1 = @$line1Read;
            push @voterStatsArray, \@values1;
            $loadCnt++;
        }
    }
    close $voterStatsFileh;
    printLine("Completed building Vote History hash for $loadCnt votes.\n");
    if ($HighVoterIDs[0] == -1) {
        printLine ("---> No High Voter ID record detected...\n");
    }
    return @voterStatsArray;
}
#---------------------------------------------------------------
#
# create the voter email binary search array
#
sub voterEmailLoad() {
    $voterEmailHeadings = "";
    open( $voterEmailFileh, $voterEmailFile )
      or die "Unable to open INPUT: $voterEmailFile Reason: $!";
    $voterEmailHeadings = <$voterEmailFileh>;
    chomp $voterEmailHeadings;
    printLine("Started Building email address array\n");

    # headings in an array to modify
    @voterEmailHeadings = split( /\s*,\s*/, $voterEmailHeadings );
    my $emailCount = 0;

    # Build the UID->survey hash
    while ( $line1Read = <$voterEmailFileh> ) {
        chomp $line1Read;
        my @values1 = split( /\s*,\s*/, $line1Read, -1 );
        push @voterEmailArray, \@values1;
        $emailCount = $emailCount + 1;
    }
    close $voterEmailFileh;
    printLine("Loaded email array: $emailCount entries");
    return @voterEmailArray;
}
#---------------------------------------------------------------
#
#  Load the configuration spreadsheet.
#  Currently it only contains the election cycle dates
#
sub load_config {
    #
    # load configuration file and load in the 20 election cycles to be used
    #
    my $dirname = dirname(__FILE__);
    $CfgFile = $dirname . "/" . $CfgFile;
    printLine("Loading Configuration from $CfgFile ...\n");
    my $bookdata = Spreadsheet::Read->new($CfgFile)
       or die "Unable to open configuration File: $CfgFile Reason: $!";
    my $MaxRows = $bookdata->[1]{maxrow};           # Save out Number of Rows in Spreadsheet
    my $MaxCols = $bookdata->[1]{maxcol};           # Save out Number of Columns in Spreadsheet
    my $DataRows = $MaxRows-1;
    #printLine("... Configuration Loaded: $DataRows Data Rows of $MaxCols columns each.\n");
    #
    # Fetch and save Excel Header row text strings into @Headings array
    #
    my $row = 0;                                                    # row index for configuration spreadsheet
    for my $j (1 .. $MaxRows) {
        @CfgHeadings = Spreadsheet::Read::row($bookdata->[1], $j);
        $row = $j;
        if (substr($CfgHeadings[0], 0, 1) eq "#") {
            next;                                                   # ignore comment lines before header row
        }
        if ($CfgHeadings[0] eq "Election Date") {                   # Found Heading line, 
            if (($CfgHeadings[1] ne "Election Type") or ($CfgHeadings[2] ne "Vote Weight")) {
                die("Invalid Configuration, Headings Not:\n Election Date, Election Type, Vote Weight\n")
            }
            last;
        }
    }
    if ($row >= $MaxRows) {
        die ("Invalid Configuration, no \"Election Date\" heading not found \n");
    }
    #
    #  Now load the election date configuration data
    #
    my ( @date, $yy, $mm, $dd, $ElecDate, $mx, $dx, $yx );          # for date conversion from spreadsheet to mm/dd/yy format
    my @ElecDates = ();                                             # electiondate array
    my $edx = 0;                                                    # electiondate index
    for my $j (($row+1) .. $MaxRows) {
        @CfgRow = Spreadsheet::Read::row($bookdata->[1], $j);       # read next config row
        if (substr($CfgRow[0], 0, 1) eq "#") {
            next;                                                   # ignore comment lines
        }
        @date = split( /[-,\/ ]/, $CfgRow[0], -1 );                 # fetch yyyy-mm-dd format date from row and split into yyyy, mm, dd in @date
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
        #
        # Verify all date fields are numbers
        #
        if (($date[$mx] !~ /^[0-9]/) or ($date[$dx] !~ /^[0-9]/) or ($date[$yx] !~ /^[0-9]/)) {
            die ("Invalid Election Date in Config at row $j \n");
        }
        $mm = sprintf( "%02d", $date[$mx] );                        # create mm, dd, yyyy as separate strings
        $dd = sprintf( "%02d", $date[$dx] );
        $yy = sprintf( "%02d", substr( $date[$yx], 2, 2 ) );
        $ElecDate = "$mm/$dd/$yy".'-'."$CfgRow[1]";    
        push (@ElecDates, $ElecDate);
        push (@electionValue, $CfgRow[2]);                          # save election voting weights
        if ($edx >= 19 ) {
            last;                                                   # only take in 20 elections
        }
        $edx++;
    }
    if ( $edx != 19 ) {
        die "Invalid Election Date Configuration, must be 20 elections defined\n";
    }
    printLine ("Configured to use these 20 elections\n");
    for my $j (0 .. 19) {
        printLine ("$ElecDates[$j] Voting Weight=$electionValue[$j]\n"); 
        $baseHeading[$j+$fixedflds] = $ElecDates[$j];            
    }
    return;
}
#
# create the precinct to politial correspondence hash
#
sub adPoliticalAll() {
    $adPoliticalHeadings = "";
    my @adPoliticalHeadings;
    
    # if no political precinct file then exit
    if ( ! (-e  $adPoliticalFile)) {
        printLine("******** Precinct XREF file $adPoliticalFile  does not exist.\n");
        printLine("******** Output Base File Will Only Contain State Races, Local Races Will Be Blank!\n");
        $Noxref = 1;
        return (-1);
    }
    printLine("Precinct XREF file is: $adPoliticalFile.\n");
    open( my $adPoliticalFileh, $adPoliticalFile )
      or die "Unable to open INPUT: $adPoliticalFile Reason: $!";
    $adPoliticalHeadings = <$adPoliticalFileh>;
    chomp $adPoliticalHeadings;
    chop $adPoliticalHeadings;

    # headings in an array to modify
    @adPoliticalHeadings = split( /\s*,\s*/, $adPoliticalHeadings );

    # Build the UID->survey hash
    while ( $line1Read = <$adPoliticalFileh> ) {
        chomp $line1Read;
        if ($line1Read eq ""){
            next;
        }
#        printLine ("line read: \"$line1Read\" \n");

        my @values1 = split( /\s*,\s*/, $line1Read, -1 );
        my $PRECINCT = $values1[0];                         # get precinct
        $adPoliticalHash{$PRECINCT} = \@values1;            # add to hash by precinct of xref data arrays
    }

    close $adPoliticalFileh;

    return @adPoliticalHash;
}
#---------------------------------------------------------------
#
# Print report line
#
sub printLine {
    my $datestring = localtime();
    ($printData) = @_;
    if ( substr( $printData, -1 ) ne "\r" ) {
        print $printFileh PROGNAME . $datestring . ' ' . $printData;
    }
    print( PROGNAME . $datestring . ' ' . $printData );
}