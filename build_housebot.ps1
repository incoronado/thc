######################################################################
### File:           build_housebot.ps1
### Author:         Mike Larson
### Description:    Primary Build Script for Housebot
###################################################################### 
Stop-Process -processname HousebotServer.exe

# Copy HB Scripts
$source="Scripts"
$dest="c:\Program Files (x86)\Housebot\Config\Scripts

$what = @("/COPYALL","/B","/MIR")
$options = @("/R:0","/W:0","/NFL","/NDL")

$cmdArgs = @("$source","$dest",$what,$options)
robocopy @cmdArgs
If($?){exit 0}

$source="HBData\HBData.mdb"
$dest="c:\Program Files (x86)\Housebot\Config\temp\HBData.mdb"


$cmdArgs = @("$source","$dest",$what,$options)
robocopy @cmdArgs
If($?){exit 0}