Option Explicit
On Error Resume Next


Dim RemoteList, i, SleepVar
RemoteList = split(GetPropertyValue("System.HBRemoteList"),",")


SleepVar = 250

Do
	For Each i in RemoteList
		UpdateStatus ("Remote-" & CStr(i))
	Next
	SongPercent 1
	SongPercent 2
	Sleep SleepVar
Loop






Function UpdateStatus (RemoteName)
	Dim NowPlayingImage, OldNotificationText, OldAlbum, OldRDSPS1, OldRDSRT1, OldRDSPS2, OldRDSRT2, OldSelectedSource
	If CInt(GetPropertyValue(RemoteName & ".Selected Source")) <> OldSelectedSource Then
	
	   OldSelectedSource = CInt(GetPropertyValue(RemoteName & ".Selected Source"))
       If CInt(GetPropertyValue(RemoteName & ".Selected Source")) = 2 Then
			SetPropertyValue RemoteName & ".Radio Station Frequency", DispFREQ(GetPropertyValue("Tuner 1 Settings.Tuner FREQ"))
	    ElseIf CInt(GetPropertyValue(RemoteName & ".Selected Source")) = 4 Then	
			SetPropertyValue RemoteName & ".Radio Station Frequency", DispFREQ(GetPropertyValue("Tuner 2 Settings.Tuner FREQ"))
		End if
	End if
	
	If CInt(GetPropertyValue(RemoteName & ".Selected Source")) = 2 Then
		If GetPropertyValue("Tuner 1 Settings.Tuner RDS PS") <> OldRDSPS1  Then
			OldRDSPS1 = GetPropertyValue("Tuner 1 Settings.Tuner RDS PS")
			SetPropertyValue RemoteName & ".Notification Text 1", OldRDSPS1
			
		Elseif (GetPropertyValue("Tuner 1 Settings.Tuner RDS RT") <> OldRDSRT1)	Then
				OldRDSRT1 = GetPropertyValue("Tuner 1 Settings.Tuner RDS RT")
		    	SetPropertyValue RemoteName & ".Notification Text 2", OldRDSRT1 
		End if
		NowPlayingImage=GetPropertyValue(RemoteName & ".Now Playing Image")
		If (NowPlayingImage <> "config\themes\THC\Radio Stations\" & GetPropertyValue("Tuner 1 Settings.Tuner FREQ") & ".png") Then
		   SetPropertyValue RemoteName & ".Now Playing Image", "config\themes\THC\Radio Stations\" & GetPropertyValue("Tuner 1 Settings.Tuner FREQ") & ".png"
		   sleep 150
		   SetPropertyValue RemoteName & ".Now Playing Image", "config\themes\THC\Radio Stations\" & GetPropertyValue("Tuner 1 Settings.Tuner FREQ") & ".png"
		End if   
	ElseIf CInt(GetPropertyValue(RemoteName & ".Selected Source")) = 4 Then
		If GetPropertyValue("Tuner 2 Settings.Tuner RDS PS") <> OldRDSPS2  Then
			OldRDSPS2 = GetPropertyValue("Tuner 2 Settings.Tuner RDS PS")
			SetPropertyValue RemoteName & ".Notification Text 1", OldRDSPS2
		Elseif (GetPropertyValue("Tuner 2 Settings.Tuner RDS RT") <> OldRDSRT2)	Then
				OldRDSRT2 = GetPropertyValue("Tuner 2 Settings.Tuner RDS RT")
		    	SetPropertyValue RemoteName & ".Notification Text 2", OldRDSRT2
		End if
		
		NowPlayingImage=GetPropertyValue(RemoteName & ".Now Playing Image")
		If (NowPlayingImage <> "config\themes\THC\Radio Stations\" & GetPropertyValue("Tuner 2 Settings.Tuner FREQ") & ".png") Then
		   SetPropertyValue RemoteName & ".Now Playing Image", "config\themes\THC\Radio Stations\" & GetPropertyValue("Tuner 2 Settings.Tuner FREQ") & ".png"
		   sleep 150
		   SetPropertyValue RemoteName & ".Now Playing Image", "config\themes\THC\Radio Stations\" & GetPropertyValue("Tuner 2 Settings.Tuner FREQ") & ".png"
		End if
	ElseIf CInt(GetPropertyValue(RemoteName & ".Selected Source")) = 1 Then
	    OldNotificationText = GetPropertyValue(RemoteName & ".Notification Text 1")
		If (OldNotificationText<> GetPropertyValue("Jukebox1.Now Playing - Title")) then
		   SetPropertyValue RemoteName & ".Notification Text 1", GetPropertyValue("Jukebox1.Now Playing - Title")
		End if
		OldNotificationText = GetPropertyValue(RemoteName & ".Notification Text 2")
		If (OldNotificationText<> GetPropertyValue("Jukebox1.Now Playing - Artist")) then
		   SetPropertyValue RemoteName & ".Notification Text 2", GetPropertyValue("Jukebox1.Now Playing - Artist")
		End if
		Sleep 150
		OldAlbum=GetPropertyValue(RemoteName & ".Jukebox - Album")
		If  (GetPropertyValue("Jukebox1.Now Playing - Album") <> OldAlbum) Then
		
		    SetPropertyValue RemoteName & ".Jukebox - Album", GetPropertyValue("Jukebox1.Now Playing - Album")  
			sleep 300
			SetPropertyValue RemoteName & ".Now Playing Image", GetPropertyValue("Jukebox1.Now Playing Image")
			sleep 300
			SetPropertyValue RemoteName & ".Now Playing Image", GetPropertyValue("Jukebox1.Now Playing Image")
			sleep 300
			SetPropertyValue RemoteName & ".Now Playing Image", GetPropertyValue("Jukebox1.Now Playing Image")
		End if
		
	ElseIf CInt(GetPropertyValue(RemoteName & ".Selected Source")) = 5 Then
	    
	    OldNotificationText = GetPropertyValue(RemoteName & ".Notification Text 1")
	    If (OldNotificationText<> GetPropertyValue("Jukebox2.Now Playing - Title")) then
		   SetPropertyValue RemoteName & ".Notification Text 1", GetPropertyValue("Jukebox2.Now Playing - Title")
		End if
		
		OldNotificationText = GetPropertyValue(RemoteName & ".Notification Text 2")
		If (OldNotificationText<> GetPropertyValue("Jukebox2.Now Playing - Artist")) then
		   SetPropertyValue RemoteName & ".Notification Text 2", GetPropertyValue("Jukebox2.Now Playing - Artist")
		End if
		Sleep 150
		OldAlbum=GetPropertyValue(RemoteName & ".Jukebox - Album")
		If  (GetPropertyValue("Jukebox2.Now Playing - Album") <> OldAlbum) Then
		    SetPropertyValue RemoteName & ".Jukebox - Album", GetPropertyValue("Jukebox2.Now Playing - Album") 
		    sleep 300
			SetPropertyValue RemoteName & ".Now Playing Image", GetPropertyValue("Jukebox2.Now Playing Image")
			sleep 300
			SetPropertyValue RemoteName & ".Now Playing Image", GetPropertyValue("Jukebox2.Now Playing Image")
			sleep 300
			SetPropertyValue RemoteName & ".Now Playing Image", GetPropertyValue("Jukebox2.Now Playing Image")
		End if
	End If
	
End Function

Sub SongPercent(ListNo)
	Dim SongLengthMin, SongLengthSec, SongLengthTotalSec, SongLengthData
	Dim SongPositionMin, SongPositionSec, SongPositionTotalSec, SongPercent, SongPositionData
    SongLengthData = Split(GetPropertyValue ("Music Player " & ListNo & ".Song Length"),":")
	SongLengthMin = CInt(SongLengthData(0))
	SongLengthSec = CInt(SongLengthData(1))
	SongLengthTotalSec = CInt((SongLengthMin*60)+SongLengthSec)
	SongPositionData = Split(GetPropertyValue ("Music Player " & CStr(ListNo) & ".Song Position"),":")
	SongPositionMin = CInt(SongPositionData(0))
   	SongPositionSec = CInt(SongPositionData(1))
	SongPositionTotalSec = CInt((SongPositionMin*60)+SongPositionSec)
	SongPercent = CInt(Round((SongPositionTotalSec/SongLengthTotalSec)*100))
	SetPropertyValue "Jukebox" & CStr(ListNo) & ".Song Percent", SongPercent
	SetPropertyValue "Jukebox" & CStr(ListNo) & ".Song Percent", SongPercent
	SetPropertyValue "Jukebox" & CStr(ListNo) & ".Song Percent", SongPercent
End Sub

   



Function DispFREQ(Frequency)
	DispFREQ = mid(Frequency,1,2) & " " & mid(Frequency,3,len(Frequency)-4) & "."  & right(Frequency,2)
End Function