var mongoose = require('mongoose');
var Schema = mongoose.Schema;

var RaceSchema = new Schema (
    {
        name: {type: String, required: true, minlength: 3, maxlength: 100}
        {
            Year; _year,
            Race: _race, 
            party: _party,
            candidate: _candidate,
                counties: [
                    {"Carson City": 6789},
                    {"Churchill": 0},
                    {"Clark": 0},
                    {"Douglas": 0},
                    {"Elko": 0},
                    {"Esmeralda": 0},
                    {"Eureka": 0},
                    {"Humbolt": 0},
                    {"Lander": 0},
                    {"Lincoln": 0},
                    {"Lyon": 0},
                    {"Mineral": 0},
                    {"Nye": 0},
                    {"Pershing": 0},
                    {"Storey": 0},
                    {"Washoe": 0},
                    {"White Pine": 0},
    }
);

//virtual for genre URL
RaceSchema
.virtual('url')
.get(function () {
    return '/catalog/genre/' + this._id
});

