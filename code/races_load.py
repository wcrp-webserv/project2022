import os
from mongoengine import *
import pandas as pd
#import models.races
import argparse
import errno
from csv  import reader
from csv import DictReader

connect('mydatabase', host='127.0.0.1', port=27017)

class Races(Document):
    race_year = StringField(required=True)
    race_name = StringField(required=True)    
    race_blob = StringField(equired=True) 
    meta = {'collection': 'races'}

#
# process command line parms
# 
parser = argparse.ArgumentParser()
parser.add_argument("directory")

args = parser.parse_args()
# echo the command line
print (args.directory)

#
# iterate through *.csv files and create mongodb documents
#
for filename in os.listdir(args.directory):
    if filename.endswith(".csv"):
                     #or filename.endswith(".py"): 

        fullname = os.path.join(args.directory, filename)
        print(fullname)

        dataFrame = pd.read_csv(fullname)
        raceblob = dataFrame.to_json()
        
        x = filename.split('.')
        racename = x[0].split('-')
        r1 = Races(race_year=racename[0], race_name=racename[1], race_blob=raceblob)
        r1.save()   

        continue
    else:
        print('File is not of type .csv')
#