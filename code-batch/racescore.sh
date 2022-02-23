#!/bin/bash
echo 'Starting RaceScoreCard Races'
echo '#=======================================================#'
echo '#     Generate Race ScoreCard            #'
echo '#=======================================================#'
export my_path="/Users/jimsievers/iCloud/my\ folders/git-c/project2022/testdata"
cd 
echo 'SD8'
python3 ~/python/racesheet.py "-d" "SD8" "-c" "CandSD8" "-y" "2022" "-s" "base-cl.csv" "-h" "data" 
echo 'SD9'
python3 ~/python/racesheet.py "-d" "SD9" "-c" "CandSD8" "-y" "2022" "-s" "base-cl.csv" "-h" "data" 
echo 'AD5'
python3 ~/python/racesheet.py "-d" "AD5" "-c" "CandAD5" "-y" "2022" "-s" "base-cl.csv" "-h" "data" 
echo 'AD21'
python3 ~/python/racesheet.py "-d" "AD21" "-c" "CandAD21" "-y" "2022" "-s" "base-cl.csv" "-h" "data" 
echo 'AD29'
python3 ~/python/racesheet.py "-d" "AD29" "-c" "CandAD29" "-y" "2022" "-s" "base-cl.csv" "-h" "data" 
echo 'AD41'
python3 ~/python/racesheet.py "-d" "AD41" "-c" "CandAD41" "-y" "2022" "-s" "base-cl.csv" "-h" "data" 
echo 'AD27'
python3 ~/python/racesheet.py "-d" "AD27" "-c" "CandAD27" "-y" "2022" "-s" "base-wa.csv" "-h" "data" 
echo 'AD30'
python3 ~/python/racesheet.py "-d" "AD30" "-c" "CandAD30" "-y" "2022" "-s" "base-wa.csv" "-h" "data" 
echo 'COMPLETED REPORTS'
