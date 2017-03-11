######################################################################
### File:           build_housebot.ps1
### Author:         Mike Larson
### Description:    Primary Build Script for Housebot
###################################################################### 
# Copy HB Scripts a
$source="Scripts"
$dest="c:\Program Files (x86)\Housebot\Config\Temp"

$what = @("/COPYALL","/B","/MIR")
$options = @("/R:0","/W:0","/NFL","/NDL")

$cmdArgs = @("$source","$dest",$what,$options)
robocopy @cmdArgs
If($?){exit 0}