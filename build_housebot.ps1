######################################################################
### File:           build_housebot.ps1
### Author:         Mike Larson
### Description:    Primary Build Script for Housebot
###################################################################### 
$password = "KKy68a?*"
pscp -pw KKy68a?* -i c:\users\mike\.ssh\id_rsa.pub Scripts/system.vbs mike@home-pc:"/Program Files (x86)/housebot/config/scripts/"

