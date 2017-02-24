######################################################################
### File:           build_housebot.ps1
### Author:         Mike Larson
### Description:    Primary Build Script for Housebot
###################################################################### 
$username = "mike"
$password = "KKy68a?*"

net use \\home-pc $password /USER:$username 

$hb_destination= "\\home-pc\Config\"
$git_source  = "Scripts\"
$logfile = $git_source + "robocopyLogFile.txt"
$switches = ("/E", "/S", "/R:0")
$roboCopyString = "*.vbs"
robocopy $git_source $hb_destination"bak" $roboCopyString $switches

