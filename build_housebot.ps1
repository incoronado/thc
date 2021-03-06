######################################################################
### File:           build_housebot.ps1
### Author:         Mike Larson
### Description:    Primary Build Script for Housebot
###################################################################### 
#Stop-Process -processname "HouseBotServer" -Force
$JOB_NAME= $env:JOB_NAME

# get HouseBotServer process
#$hb = Get-Process HouseBotServer -ErrorAction SilentlyContinue
#if ($hb) {
  # try gracefully first
#  $hb.CloseMainWindow()
  # kill after five seconds
#  Start-Sleep -s 5
#  if (!$hb.HasExited) {
#    $hb | Stop-Process -Force
#  }
#}
Remove-Variable hb

#$lockfile= "C:\Program Files (x86)\Housebot\Config\HBData.ldb"
#if (Test-Path $lockfile) {
#  Remove-Item $lockfile
#}


# Copy HB Scripts
$source="Scripts"
$dest="c:\Program Files (x86)\Housebot\Config\Scripts"

$what = @("/COPYALL","/B","/MIR", "/SEC")
$options = @("/R:0","/W:0","/NFL","/NDL")

$cmdArgs = @("$source","$dest",$what,$options)
robocopy @cmdArgs


#$source="Themes"
#$dest="c:\Program Files (x86)\Housebot\Config\Themes"

#$what = @("/COPYALL","/B","/MIR")
#$options = @("/R:0","/W:0","/NFL","/NDL")

#$cmdArgs = @("$source","$dest",$what,$options)
# robocopy @cmdArgs


If ($JOB_NAME -eq "Development") {
    $source="HBData"
    $dest="c:\Program Files (x86)\Housebot\Config"
    $filename="HBData.mdb"
    $cmdArgs = @("$source","$dest",$filename,$options)
    robocopy @cmdArgs
}

$LASTEXITCODE = 0


exit $LastExitCode