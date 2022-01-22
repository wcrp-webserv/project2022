#!/bin/bash
echo 'Starting i360 Survey Log Analyzer'
echo '#=======================================================#'
echo '#     This script analyzers i360 survey logs            #'
echo '#=======================================================#'
cd ~/Downloads
pwd
python ~/python/nvreport.py -infile i360 -datadir ~/Dropbox/2022-electioneering/p1-canvass/3survey-responses/ -rptfile SurveyResponseRpt
echo 'COMPLETED REPORT'
