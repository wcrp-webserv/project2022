#!/bin/bash
echo 'Starting voter list build'
echo '#=======================================================#'
echo '#     This script transforms NVSOS voter downloads      #'
echo '#     Transforming $COUNTY                              #'
echo '#=======================================================#'
perl ~/perl5/nvvoter1.pl -infile *VtHst*
echo 'startting sort one'
echo '#=======================================================#'
echo '#     Sorting intermediate file                         #'
echo '#=======================================================#'
csvsort -c 1 -e Latin1 voterdata.csv > voterdata-s.csv
echo  'completed sort one'
echo '#=======================================================#'
echo '#     Final step in transformation                      #'
echo '#=======================================================#'
perl ~/perl5/nvvoter2.pl -infile *ElgbVtr*
echo '#=======================================================#'
echo '#     Cleaning up after transformation                  #'
echo '#=======================================================#'
