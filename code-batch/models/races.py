import mongoengine

class Races(Document):
    race_year = StringField(required=True)
    race_name = StringField(required=True)    
    race_blob = StringField(equired=True) 
    meta = {'collection': 'races'}