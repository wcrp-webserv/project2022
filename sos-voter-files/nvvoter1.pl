#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
# wcrp-voter-nvhistory
#  -- nvvoter1
#  Convert the NVSOS voter data to voter Statistic lines
#            VoterList.VtHst
#
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
#use strict;
use warnings;
$| = 1;
use File::Basename;
use DBI;
use Data::Dumper;
use Getopt::Long qw(GetOptions);
use Time::Piece;
use Time::Seconds;
use Math::Round;
use Text::CSV qw( csv );
use Spreadsheet::Read;
use constant PROGNAME => "NVVOTER1 - ";

no warnings "uninitialized";

=head1 Function
=over
=head2 Overview

=cut

my $records;

#
#  Configuration SpreadSheet File Name, Header & Data Row Arrays
#
my $CfgFile = "nvconfig.xlsx";                  # program configuration spreadsheet
my @CfgHeadings =();                            # Array of Text Headings for spreadsheet
my @CfgRow =();                                 # Data from the Row of spreadsheet currently being processed

my $voterHistoryFile = "VoterList.VtHst.43842.060420175555.csv";
my $voterHistoryFileh;
my @voterHistoryLine = ();
my %voterHistoryLine;

my $voterDataHeading = "";
my $voterDataFile    = "voterdata.csv";
my $voterDataFileh;
my @voterData;
my %voterDataLine = ();
my @voterDataLine;

my @electionValue = ();

my $printFile = "print.txt";
my $printFileh;

my $helpReq   = 0;
my $fileCount = 0;

my $csvHeadings = "";
my @csvHeadings;
my $line1Read = '';
my $linesRead = 0;
my $printData;
my $linesWritten = 0;
my $maxFiles;
my $maxLines;
my $csvRowHash;
my @csvRowHash;
my %csvRowHash   = ();
my $stateVoterID = 0;
my @date;
my $adjustedDate;
my $before;
my $vote;
my $cycle;
my $totalVotes      = 0;
my $linesIncRead    = 0;
my $linesIncWritten = 0;
my $ignored         = 0;
my $currentVoter;

my @voterDataHeading = (
    "statevoterid",
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
    "Score",           #29 Calculated
);


my @precinctPolitical;
my $RegisteredDays   = 0;
my $pollCount        = 0;
my $absenteeCount    = 0;
my $provisionalCount = 0;
my $earlyCount       = 0;
my $activeVOTERS     = 0;
my $activeREP        = 0;
my $activeDEM        = 0;
my $activeOTHR       = 0;
my $totalVOTERS      = 0;
my $totalGENERALS    = 0;
my $totalPRIMARIES   = 0;
my $totalPOLLS       = 0;
my $totalABSENTEE    = 0;
my $totalPROVISIONAL = 0;
my $totalMAIL        = 0;
my $totalSTR         = 0;
my $totalMOD         = 0;
my $totalWEAK        = 0;
my $votesTotal       = 0;
my $voterScore       = 0;
my $voterScore2      = 0;

my @voterHeadingDates =
  ( 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 );
my @voterEarlyDates =
  ( 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 );

#
#  Array to compile the highest voter ID that voted in each of the 20 elections being tracked
#
my @HighVoterID = (
    0,                                          # VoterID = 0 indicates this record
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0     # Highest Voter ID that voted in the 20 elections (inits at 0)
);

#
# main program controller
#
sub main {

    # Open file for messages and errors
    open( $printFileh, ">", "$printFile" )
      or die "Unable to open PRINT: $printFile Reason: $!";


    # Parse any parameters
    GetOptions(
        'infile=s'   => \$voterHistoryFile,
        'outfile=s'  => \$voterDataFile,
        'config=s'   => \$CfgFile,
        'maxlines=n' => \$maxLines,
        'maxfiles=n' => \$maxFiles,
        'help!'      => \$helpReq,

    ) or die "Incorrect usage! \n";

    if ($helpReq) {
        print "Come on, it's really not that hard. \n";
        die;
    }

    load_config();                                                  # load configuration spreadsheet
    #
    # Open Secretary of State Vote History .csv File
    #
    printLine("My Vote History File is: $voterHistoryFile. \n");
    open ($voterHistoryFileh, $voterHistoryFile) 
        or die ("Unable to open Vote History File: $voterHistoryFile Reason: $! \n");
    #
    # prepare to use text::csv module
    # build the constructor
    my $csv = Text::CSV->new(
        {
            binary             => 1,  # Allow special character. Always set this
            auto_diag          => 1,  # Report irregularities immediately
            allow_whitespace   => 0,
            allow_loose_quotes => 1,
            quote_space        => 0,
        }
    );
    @csvHeadings = $csv->header($voterHistoryFileh);

    # on input these column headers contained a space - replace headers
    $csvHeadings[0] = "uniquevoteid";
    $csvHeadings[1] = "voterid";
    $csvHeadings[2] = "electiondate";
    $csvHeadings[3] = "votecode";
    $csv->column_names(@csvHeadings);

    # Build heading for new voting record, open output file and write header row
    #
    $voterDataHeading = join( ",", @voterDataHeading );
    $voterDataHeading = $voterDataHeading . "\n";
    open( $voterDataFileh, ">$voterDataFile" )
      or die "Unable to open voter info file: $voterDataFile Reason: $! \n";
    print $voterDataFileh $voterDataHeading;

##
    #
    #  initialize oldest election date we care about
    #
    my $string         = $voterDataHeading[20];                                 # Fetch Oldest Election we're configured for
    my $oldestElection = substr( $string, 0, 8 );                               # extract date
    my $oldestDate     = Time::Piece->strptime( $oldestElection, "%m/%d/%y" );  # Convert to Date/Time object
    printLine("Oldest Election Date: $oldestElection\n");                       # display to logging stream(s)
    #
    # initialize binary election date/time object arrays from configuration test dates
    #
    for ( $vote = 1 ; $vote <= 20 ; $vote++ ) {
        my $edate        = substr( $voterDataHeading[$vote], 0, 8 );
        my $electiondate = Time::Piece->strptime( $edate, "%m/%d/%y" );
        $voterHeadingDates[$vote] = $electiondate;                              # this is election date
        $voterEarlyDates[$vote]   = ( $electiondate - ONE_WEEK ) - ONE_WEEK;    # this is early voting start
    }
    #
    #  At this point:
    #     1. $voterDataHeading[1-20]  contain text election date & type
    #     2. $voterHeadingDates[1-20] contain Date/Time object election dates
    #     3. $voterEarlyDates[1-20]   contain Date/Time object early voting start dates
    #
    # Initialize process loop and open first output
    $linesRead = 0;

#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
# main process loop.
#  initialize program
#
#  for each of several vote records for same voter in voterHistoryFileh convert to a single row for unique voter
#    - currentVoter = record-id
#    - get record from voterHistoryFileh
#    - if currentVoter is same as stateVoterID then add segment to row
#      else create calculated values and write-the-row to output file,
#      stateVoterID = currentVoter
#   endloop
#
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

  NEW:
    #
    #  Read Voter History File row by row (VoterHistoryID,VoterID,ElectionDate,VoteType)
    #
    while ( $line1Read = $csv->getline_hr($voterHistoryFileh) ) {
        $linesRead++;
        $linesIncRead += 1;
        #       if ( $linesIncRead >= int( 10000) ) {
        #           printLine("$linesRead lines read \n");      # log progress
        #           $linesIncRead = 0;
        #       }
        %csvRowHash = %{$line1Read};                # Hash this row to row titles
        $currentVoter = $csvRowHash{"voterid"};     # get voter id of this vote record
        if ( $stateVoterID == 0 ) {
            # This is the very first record of the history file.
            # Set this voter ID as the StateVoterID being processed
            $stateVoterID  = $currentVoter;
            %voterDataLine = ();                    # initialize voter data line we're building
            # clear all election data buckets to blanks
            for ( $cycle = 0 ; $cycle <= 20 ; $cycle++ ) {
                $voterDataLine{ $voterDataHeading[$cycle] } = " ";
            }
        }
        # for all records build a line for each voter with all their
        # votes by election
      next_voter:
        if ( $currentVoter eq $stateVoterID ) {
            #
            #  Processing the next vote record for this voter ID
            #  Set ID in output row in case this is 1st record for this voter
            #
            $voterDataLine{"statevoterid"} = $csvRowHash{"voterid"};
            #
            # place vote in correct election bucket (14 days <= electiondate)
            #
            my $vdate;
            my $votedate = substr( $csvRowHash{"electiondate"}, 0, 10 );        # fetch election date from SOS record
            my $datelen =length($votedate);
            #
            #  Convert SOS Vote History date to Date/Time object
            if ($datelen <= 7) {
                $vdate    = Time::Piece->strptime( $votedate, "%m/%d/%y" );  #change %Y -> %y??
            } else {
                $vdate    = Time::Piece->strptime( $votedate, "%m/%d/%Y" );  #change %Y -> %y??
            }
            if ( $vdate < $oldestDate ) {

                # ignore records for elections older than we are looking for
                $ignored += 1;
                next;
            }
            my $baddate = 1;

            # find the correct election for this vote
            # dates must be in Time::Piece format
            for ( $cycle = 1, $vote = 1 ; $cycle <= 20 ; $cycle++, $vote += 1 )
            {

                # create the earlydate used for testing the votedate
                my $electiondate  = $voterHeadingDates[$vote];
                my $twoweeksearly = $voterEarlyDates[$vote];
                my $nowdate       = $twoweeksearly->mdy;

                # test to find if the votedate fits a slot, add the vote
                if ( $vdate >= $twoweeksearly && $vdate <= $electiondate ) {
                    #
                    #  This is the election this vote is for, stash the code
                    $voterDataLine{ $voterDataHeading[$vote] } = $csvRowHash{"votecode"};

                    # add to total votes for this voter & say not a bad election date
                    $totalVotes++;
                    $baddate = 0;
                    #
                    # See if voter ID is higher that current highest for this election
                    #
                    if ( $currentVoter > $HighVoterID[$vote]) {
                        $HighVoterID[$vote] = $currentVoter;                # this is now highest voter ID in this election
                    }
                    last;
                }
            }
            #if ( $baddate != 0 ) {
            #    printLine("Unknown Election Date  $csvRowHash{'electiondate'}  for voter $csvRowHash{'voterid'} \n");
            #}
            next;
        }
        else {
            if ( $voterDataLine[0] eq " " ) {
                next; # we ignored all records for a voter, don't write anything
            }
    #
    #  End of compilation for this voter, add calculated values to output record
    #
    # Calculate rest of the values for this voter
    #
            evaluateVoter();
            #
            # put caclulated values in output data line
            #
            $voterDataLine{ $voterDataHeading[21] } = $votesTotal;
            $voterDataLine{ $voterDataHeading[22] } = $generalCount;
            $voterDataLine{ $voterDataHeading[23] } = $primaryCount;
            $voterDataLine{ $voterDataHeading[24] } = $pollCount;
            $voterDataLine{ $voterDataHeading[25] } = $absenteeCount;
            $voterDataLine{ $voterDataHeading[26] } = $earlyCount;
            $voterDataLine{ $voterDataHeading[27] } = $provisionalCount;
            $voterDataLine{ $voterDataHeading[28] } = $voterRank;
            $voterDataLine{ $voterDataHeading[29] } = $voterScore2;
            #
            # prepare to write out the voter data
            @voterData = ();
            foreach (@voterDataHeading) {
                push( @voterData, $voterDataLine{$_} );
            }
            print $voterDataFileh join( ',', @voterData ), "\n";
            %voterDataLine = ();
            $linesWritten++;
            $linesIncWritten++;
            $totalVotes = 0;
            $linesRead++;
            if ( $linesIncWritten == 2000 ) {
                printLine("$linesWritten lines written \r");
                $linesIncWritten = 0;
            }

            # clear output to blanks for next voter
            for ( $cycle = 0 ; $cycle <= 20 ; $cycle++ ) {
                $voterDataLine{ $voterDataHeading[$cycle] } = " ";
            }

            # process this input record which is for the next voter
            $stateVoterID = $currentVoter;
            goto next_voter;
        }

        #
        # For now this is the in-elegant way I detect completion
        if ( eof(voterHistoryFileh) ) {
            goto EXIT;
        }
        next;
    }
}

#
# call main program controller
main();
#
# Common Exit
EXIT:

Add_HighIDs();                              # write voter ID 0 record with highest ID in each election
close(voterHistoryFileh);
close($voterDataFileh);

printLine("<===> Completed processing of: $voterHistoryFile \n");
printLine("<===> Total Records Read: $linesRead \n");
printLine("<===> Total Records written: $linesWritten \n");
printLine("<===> Total Old Vote Records Ignored: $ignored \n");

close($printFileh);
exit;

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
sub evaluateVoter {
    my $generalPollCount  = 0;
    my $generalEarlyCount = 0;
    my $generalNotVote    = 0;
    my $notElegible       = 0;
    my $primaryPollCount  = 0;
    my $primaryEarlyCount = 0;
    my $primaryNotVote    = 0;
    my $badcode           = 0;
    my $badstring         = "";
    my $oldestCast        = 0;
    $generalCount     = 0;
    $primaryCount     = 0;
    $pollCount        = 0;
    $absenteeCount    = 0;
    $earlyCount       = 0;
    $provisionalCount = 0;
    $votesTotal       = 0;
    $voterRank        = '';

    #set pointer to first vote in list
    my $vote = 1;

    for ( my $cycle = 1 ; $cycle <= 20 ; $cycle++, $vote += 1 ) {
        $badcode   = 1;
        $badstring = ( $voterDataLine{ $voterDataHeading[$vote] } );

# each election type is specified with its date - we only process primary/general
# skip mock election
        if ( ( $voterDataHeading[$vote] ) =~ m/mock/ ) {
            $badcode = 0;
            next;
        }

        # skip special election
        if ( ( $voterDataHeading[$vote] ) =~ m/special/ ) {
            $badcode = 0;
            next;
        }

        #skip sparks election
        if ( ( $voterDataHeading[$vote] ) =~ m/sparks/ ) {
            $badcode = 0;
            next;
        }
        #
        # record a general vote
        # if there is no vote recorded shown with a "blank" then NOT ELEGIBLE
        #
        if ( ( $voterDataHeading[$vote] ) =~ m/general/ ) {
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq ' ' ) {
                $badcode = 0;
                $notElegible += 1;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq "" ) {
                $badcode = 0;
                $notElegible += 1;
                next;
            }
            #
            # the following vote codes are supported
            # - EV early vote
            # - FW federal write in
            # - MB mail ballot
            # - PP polling place
            # - PV provisional vote
            # - BR ballot received (prior to election day, becomes MB at election time)
            #
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'N' ) {
                $badcode = 0;
                $generalNotVote += 1;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'PP' ) {
                $generalPollCount += 1;
                $generalCount     += 1;
                $pollCount        += 1;
                $votesTotal       += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'FW' ) {
                $generalPollCount += 1;
                $generalCount     += 1;
                $pollCount        += 1;
                $votesTotal       += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'EV' ) {
                $generalEarlyCount += 1;
                $earlyCount        += 1;
                $generalCount      += 1;
                $votesTotal        += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'MB' ) {
                $generalEarlyCount += 1;
                $generalCount      += 1;
                $earlyCount        += 1;
                $absenteeCount     += 1;
                $votesTotal        += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }
             if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'PV' ) {
                $generalCount      += 1;
                $provisionalCount  += 1;
                $votesTotal        += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }

            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'BR' ) {
              #  $generalCount     += 1;
              # $provisionalCount += 1;
              #  $votesTotal       += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }
            if ( $badcode != 0 ) {
                printLine(
"Unknown General Election Code $badstring for voter $currentVoter \n"
                );
                $badcode = 0;
            }
        }

        # record a primary vote
        # if there is no vote recorded shown with a "blank" then NOT ELEGIBLE
        #
        if ( ( $voterDataHeading[$vote] ) =~ m/primary/ ) {
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq ' ' ) {
                $notElegible += 1;
                $badcode = 0;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq "" ) {
                $notElegible += 1;
                $badcode = 0;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'N' ) {
                $primaryNotVote += 1;
                $badcode = 0;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'PP' ) {
                $primaryPollCount += 1;
                $primaryCount     += 1;
                $pollCount        += 1;
                $votesTotal       += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'EV' ) {
                $primaryEarlyCount += 1;
                $earlyCount        += 1;
                $primaryCount      += 1;
                $votesTotal        += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'MB' ) {
                $primaryEarlyCount += 1;
                $primaryCount      += 1;
                $earlyCount        += 1;
                $absenteeCount     += 1;
                $votesTotal        += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }
            if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'BR' ) {
               # $primaryEarlyCount += 1;
               # $primaryCount      += 1;
               # $earlyCount        += 1;
               # $absenteeCount     += 1;
               # $votesTotal        += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }if ( $voterDataLine{ $voterDataHeading[$vote] } eq 'PV' ) {
                $primaryCount     += 1;
                $provisionalCount += 1;
                $votesTotal       += 1;
                $oldestCast = $vote;
                $badcode    = 0;
                next;
            }
            if ( $badcode != 0 ) {
                printLine(
"Unknown Primary Election Code $badstring for voter $currentVoter } \n"
                );
                $badcode = 0;
            }
        }
        if ( $badcode != 0 ) {
            printLine(
                "Unknown Vote Code $badstring for voter $currentVoter \n");
            $badcode = 0;
        }
    }

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
    ## if registered > 8 years   gen >= 6 && gen <= 12 and pri >= 0      = STRONG

    if ( $votesTotal > 0 ) {
        $voterScore  = ( $generalCount + $primaryCount ) / ($oldestCast) * 10;
        $voterScore2 = round($voterScore);
    }else{
        $voterScore2 = 0;                               # if never voted, call WEAK
    }

    if ( $voterScore2 > 5 ) {
        $voterRank = "STRONG";
    }
    if ( $voterScore2 >= 3 && $voterScore2 <= 5 ) {
        $voterRank = "MODERATE";
    }
    if ( $voterScore2 < 3 ) {
        $voterRank = "WEAK";
    }

    #
    # Set voter strength rating
    #
    if    ( $voterRank eq 'STRONG' )   { $totalSTR++; }
    elsif ( $voterRank eq 'MODERATE' ) { $totalMOD++; }
    elsif ( $voterRank eq 'WEAK' )     { $totalWEAK++; }

    $totalGENERALS    = $totalGENERALS + $generalCount;
    $totalPRIMARIES   = $totalPRIMARIES + $primaryCount;
    $totalPOLLS       = $totalPOLLS + $pollCount;
    $totalABSENTEE    = $totalABSENTEE + $absenteeCount;
    $totalPROVISIONAL = $totalPROVISIONAL + $provisionalCount;
}
#---------------------------------------------------------------
#
#  Load the configuration spreadsheet.
#  Currently it only contains the election cycle dates, types and vote weights
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
        $ElecDate = "$mm/$dd/$yy $CfgRow[1]";                       # build "mm/dd/yy type" string
        push (@ElecDates, $ElecDate);                               # save election column headers
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
        printLine ("$ElecDates[$j] Voting Weight=$electionValue[$j]\n");                            # display on console and in print file
        $voterDataHeading[$j+1] = $ElecDates[$j];                   #copy to active header
    }
    return;
}

#----------------------------------------------------------------
#
#  Write thesummary record at EOF listing highest ID that voted in each election
#
sub Add_HighIDs {
    #
    # prepare to write out the voter data
    # build array of zeroes to start
    #
    @voterData = ();
    foreach (@voterDataHeading) {
        push( @voterData, 0 );
    }
    for my $i (1 .. 20) {
        $voterData[$i] = $HighVoterID[$i];                      # add in highest voter IDs for each election
    }
    print $voterDataFileh join( ',', @voterData ), "\n";
    return;
}
#---------------------------------------------------------------
#
# Print report line to screen
#
sub printLine {
    my $datestring = localtime();
    ($printData) = @_;
    if ( substr( $printData, -1 ) ne "\r" ) {
        print $printFileh PROGNAME . $datestring . ' ' . $printData;
    }
    print( PROGNAME . $datestring . ' ' . $printData );
    return;
}
