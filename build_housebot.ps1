######################################################################
### File:           build_housebot.ps1
### Author:         Mike Larson
### Description:    Primary Build Script for Housebot
###################################################################### 

$hb_destination= "\\home-pc\Config\"
$git_source  = "Scripts\"
$logfile = $git_source + "robocopyLogFile.txt"
$switches = ("/E", "/S", "/R:0", "/LOG:$logfile")
$roboCopyString = "*.vbs"
robocopy $git_source $hb_destination"bak" $roboCopyString $switches

