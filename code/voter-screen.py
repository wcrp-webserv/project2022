#mport numpy as np
import argparse
import errno
from csv  import reader
from csv import DictReader

#
# get command line arguments
#
parser = argparse.ArgumentParser()
parser.add_argument("infile")
parser.add_argument("type")
parser.add_argument("district")
parser.add_argument("-ad", "--AD")
parser.add_argument("-sd", "--SD")

args = parser.parse_args()
# echo the command line
print (args.infile, args.type, args.district)

column = ""
if args.type == "SD":
    column = "SenDist"
elif args.type == "AD":
    column = "AssmDist"
analyze_file = args.infile

# read data source 
with open(analyze_file, mode='r') as csv_file:
    csv_reader = DictReader(csv_file)
    print ("opened file: {}".format(analyze_file))


# variable definitions
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


    voted = ['BR', 'EV', 'MB', 'PP', 'PV']

    # create document
    for row in csv_reader:
        if (row[column] != args.district):
            continue
        allvoters+=1
        vote = str(row['11/03/20 general'])

        if (row['Party'] == 'Democrat'):
            democrats = democrats+1
            if vote in voted:
                anyvote = anyvote+1
                demvote = demvote+1

        elif (row['Party'] == 'Republican'):
            republicans = republicans+1
            if vote in voted:
                anyvote = anyvote+1
                repvote = repvote+1
        elif (row['Party'] == 'Non-Partisan'):
            nonps = nonps + 1
            if vote in voted:
                anyvote = anyvote+1
                nonpvote = nonpvote+1
        elif (row['Party'] == 'Independent American Party'):
            iaps = iaps + 1
            if vote in voted:
                anyvote = anyvote+1
                iapvote = iapvote+1
        else:
            other = other + 1
            if vote in voted:
                anyvote = anyvote+1
                othvote = othvote +1
#  these values will be written into a table for report, for now just because
#  VOTERS: total active voters by party group
#   VOTED: total votes cast by party group
# TURNOUT: turnout percentage by group

    print( 'VOTERS: {:6d} REP: {:6d} DEM: {:6d} NONP: {:6d} IAP: {:6d} OTH: {:6d}'.format(allvoters, republicans, democrats, nonps, iaps, other))
    print( ' VOTED: {:6d} REP: {:6d} DEM: {:6d} NONP: {:6d} IAP: {:6d} OTH: {:6d}'.format(anyvote, repvote, demvote, nonpvote, iapvote, othvote))
    print( '   T/O:  {:0.3f} REP:  {:0.3f} DEM:  {:0.3f} NONP:  {:0.3f} IAP:  {:0.3f} OTH:  {:0.3f}'.format(anyvote/allvoters, repvote/republicans, demvote/democrats, nonpvote/nonps, iapvote/iaps, othvote/other))
    

    # count by party 
    
        