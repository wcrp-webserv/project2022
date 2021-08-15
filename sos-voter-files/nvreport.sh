#!/bin/bash
echo 'Starting i360 Survey Log Analyzer'
echo '#=======================================================#'
echo '#     This script analyzers i360 survey logs            #'
echo '#=======================================================#'
export my_path=/Users/jimsievers/symlinks/2022-electioneering/canvass/3survey-responses/surveys
cd ~/Downloads
echo $my_path
python ~/python/nvreport.py -infile i360 -datadir ~/symlinks/electioneering/2022-electioneering/canvass/3survey-responses/ -rptfile SurveyResponseRpt
echo 'COMPLETED REPORT'