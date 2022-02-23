#!/bin/bash
echo 'Starting creating voter history'
echo '#=======================================================#'
echo '#     This script transforms NVSOS voter downloads      #'
echo '#     data comes in two files:                          #'
echo '#     - data input is in two files from NVSOS download: #'
echo '#       *VtHst* ---- voter history 20 past votes        #'
echo '#       *ElgbVtr* -- elegible voters info               #'
echo '#     data is output into three files                   #'
echo '#        base.csv     base data output                  #'
echo '#        precinct.csv predinct summaryies for disctrict #'
echo '#        print.txt    printed log                       #'
echo '#=======================================================#'

perl ~/perl5/nvvoter1.pl -infile *VtHst*

echo 'sorting voter history'
echo '#=======================================================#'
echo '#     Sorting intermediate file                         #'
echo '#=======================================================#'

csvsort -c 1 -e Latin1 voterdata.csv > voterdata-s.csv

echo 'completed sorting voter history'
echo '#=======================================================#'
echo '#     Final step in transformation                      #'
echo '#=======================================================#'

echo 'starting creating of base file'

perl ~/perl5/nvvoter2.pl -infile *ElgbVtr*

echo 'completed creating base file'


echo '#=======================================================#'
echo '#     Cleaning up after transformation                  #'
echo '#=======================================================#'
