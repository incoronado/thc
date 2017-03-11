######################################################################
### File:           build_housebot.ps1
### Author:         Mike Larson
### Description:    Primary Build Script for Housebot
###################################################################### 
Stop-Process -processname HousebotServer.exe

# Copy HB Scripts
$source="Scripts"
$dest="c:\Program Files (x86)\Housebot\Config\Scripts"

$what = @("/COPYALL","/B","/MIR")
$options = @("/R:0","/W:0","/NFL","/NDL")

$cmdArgs = @("$source","$dest",$what,$options)
robocopy @cmdArgs

$source="HBData"
$dest="c:\Program Files (x86)\Housebot\Config\temp"
$filename="HBData.mdb"

$cmdArgs = @("$source","$dest",$filename,$what,$options)
robocopy @cmdArgs
If($?){exit 0}