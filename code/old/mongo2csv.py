from mongoengine import *
import numpy as np
import csv
import mongoengine

connect('mydatabase', host='127.0.0.1', port=27017)

class CountyVote(EmbeddedDocument):
    county_name = StringField(max_length=50)
    county_votes = StringField(max_length=6)

class Candidate(EmbeddedDocument):
    party = StringField(required=True, max_length=3)
    candidate = StringField(required=True, max_length=50)
    counties = ListField(EmbeddedDocumentField(CountyVote))
    meta = {'allow_inheritance': True}   

class Races(Document):
    race_year = StringField(required=True)
    race_name = StringField(required=True)    
    candidates = ListField(EmbeddedDocumentField(Candidate)) 
    meta = {'allow_inheritance': True}   

# read record from database
# 
race_yr = Races.objects(race_year='2020',race_name='sd6')
 


# add candidate and party then votes by county
row_index = 0

# create iterators
row_headers = result[0]
row_candidates = result[1]
list_headers = row_headers[3:]
list_candidates = row_candidates[3:]
list_counties = result[-17:]

# create document
r1 = Races(race_year=result[1,0], race_name=result[1,1])
r1.save()   

for item_candidate in list_candidates:
    c1 = Candidate(party=list_headers[row_index], candidate=item_candidate)
    r1.candidates.append(c1)
    print('append a candidate {} a {}'.format(item_candidate, list_headers[row_index])) 
    
    for item_county in list_counties:
        v1 = CountyVote(county_name = item_county[2], county_votes = item_county[3])
        c1.counties.append(v1)
        # print('append a county {} with votes {}'.format(item_county[2], item_county[3]))
    row_index += 1
    r1.save()

        