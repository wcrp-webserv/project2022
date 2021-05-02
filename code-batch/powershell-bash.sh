# Get Version number and date of Downloaded files
#
echo  "Complete the file name VoterList.ElgbVtr."
#$vname = Read-Host
if ( $vname ) {

}else{
   $vname="45099.073019143713.csv"
}
$elgb="VoterList.ElgbVtr.$vname"
$hist="VoterList.VtHst.$vname"
#
echo "File Names are $elgb, and $hist"
#   ===========>Input Sec of State Voter History File
#   ===========>Output voterdata.csv
$dt = Get-Date -UFormat %c
Write-Host "PsScript - $dt > perl nvvoter1-new.pl -infile $hist"
perl nvvoter1-new.pl -infile $hist
if ( $LastExitCode ) {
   Write-Host "Fatal Error -- Aborting..."
   Read-Host -Prompt "Press Enter to exit"
   Break
}
#
# Sort voterdata.csv file into voterdata-s.csv file
#
$dt = Get-Date -UFormat %c
Write-Host "PsScript - $dt - Sorting voterdata.csv into voterdata-s.csv quoted file..."
Add-Content print.txt "PsScript - $dt - Sorting voterdata.csv into voterdata-s.csv quoted file..."
Import-Csv .\voterdata.csv | sort  {[int]$_.statevoterid} | Export-Csv -Path .\voterdata-s.csv -NoTypeInformation
#
# ============> Input Sec of State Eligible Voters File
# ============> voterdata.csv
# ============> Optional email add file
# ============> Output base.csv Zoho Load File
#
$dt = Get-Date -UFormat %c
Write-Host "PsScript - $dt > perl nvvoter1-new.pl -infile $elgb"
perl nvvoter2-new.pl -infile $elgb
if ( $LastExitCode ) {
   Write-Host "Fatal Error -- Aborting..."
   Read-Host -Prompt "Press Enter to exit"
   Break
}
#
#  Keep Window Open until key hit
# 
$dt = Get-Date -UFormat %c
Write-Host "PsScript - $dt - File base.csv Ready for Uploading to Zoho`n"
Read-Host -Prompt "Press Enter to exit"