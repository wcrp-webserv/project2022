#!/bin/bash
echo 'Starting creating voter history'
echo '#==========================================================#'
echo '#     This script transforms NVSOS voter downloads         #'
echo '#     data comes in two files:                             #'
echo '#     - data input in serveral files from NVSOS download:  #'
echo '#     -config        nvconfig.xlsx                         #'
echo '#     -infile        *VtHst*      vote history             #'
echo '#     -regfile       *ElgbVtr*    elegible voters info     #'
echo '#     -xref          input precinct to district xref       #'
echo '#     -outfile       base.csv     base data out            #'
echo '#     -emailfile     emails.csv   input emails             #'
echo '#     other output files                                   #'
echo '#                                                          #'
echo '#        precinct.csv predinct summary for disctrict       #'
echo '#        print.txt    printed log                          #'
echo '#==========================================================#'

python ~/python/nvvoter.py -config ~/python/nvconfig.xlsx -infile *VtHst* -regfile *ElgbVtr* -outfile base.csv -xref PreinctXref.xlsx

echo 'completed creating base file'

echo '#=======================================================#'
echo '#     Cleaning up after transformation                  #'
echo '#=======================================================#'
