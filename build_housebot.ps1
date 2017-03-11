######################################################################
### File:           build_housebot.ps1
### Author:         Mike Larson
### Description:    Primary Build Script for Housebot
###################################################################### 
Stop-Process -processname "HouseBotServer" -Force
Remove-Item "C:\Program Files (x86)\Housebot\Config\HBData.ldb"

# Copy HB Scripts
$source="Scripts"
$dest="c:\Program Files (x86)\Housebot\Config\Scripts"

$what = @("/COPYALL","/B","/MIR")
$options = @("/R:0","/W:0","/NFL","/NDL")

$cmdArgs = @("$source","$dest",$what,$options)
robocopy @cmdArgs

$source="Themes"
$dest="c:\Program Files (x86)\Housebot\Config\Themes"

$what = @("/COPYALL","/B","/MIR")
$options = @("/R:0","/W:0","/NFL","/NDL")

$cmdArgs = @("$source","$dest",$what,$options)
robocopy @cmdArgs




$source="HBData"
$dest="c:\Program Files (x86)\Housebot\Config"
$filename="HBData.mdb"

$cmdArgs = @("$source","$dest",$filename,$options)
robocopy @cmdArgs
If($?){exit 0}

& "C:\Program Files (x86)\Housebot\HouseBotServer.exe"