import pymongo

myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["mydatabase"]

mycol = mydb["races"]

#mydict = { "name": "John", "address": "Highway 37" }

mydict= {
            "year": "2020",
            "race": "SD6", 
            "party": "DEM",
            "candidate": "Bob Dolan",
                "counties": [
                    {"Carson City": "6789", "0", "0"."0", "0","0"}},
                    {"Churchill": {"0", "0", "0"."0", "0","0"}},
                    {"Clark":  {"0", "0", "0"."0", "0","0"}},
                    {"Douglas": "0"},
                    {"Elko": "0"},
                    {"Esmeralda": "0"},
                    {"Eureka": "0"},
                    {"Humbolt": "0"},
                    {"Lander": "0"},
                    {"Lincoln": "0"},
                    {"Lyon": "0"},
                    {"Mineral": "0"},
                    {"Nye": "0"},
                    {"Pershing": "0"},
                    {"Storey": "0"},
                    {"Washoe": "0"},
                    {"White Pine": "0"},
                ]
            }



x = mycol.insert_one(mydict)
print(x.inserted_id)

    



