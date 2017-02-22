Dim SongLengthMin
Dim SongLengthSec
Dim SongLengthTotalSec
Dim SongPositionMin
Dim SongPositionSec
Dim SongPositionTotalSec
Dim SongPercent
Dim Status

' Updating every 2 Seconds is reasonable
SleepVar = 2000

Status = ""
On Error Resume Next

Do
   Sleep 5
   SongLengthSec = Mid(GetPropertyValue("Music Player.Song Length"),4,2)
   SongLengthMin = Mid(GetPropertyValue("Music Player.Song Length"),1,2)
   SongLengthTotalSec = (SongLengthMin*60)+SongLengthSec
   SongPositionSec = Mid(GetPropertyValue("Music Player.Song Position"),4,2)
   SongPositionMin = Mid(GetPropertyValue("Music Player.Song Position"),1,2)
   SongPositionTotalSec = (SongPositionMin*60)+SongPositionSec
   SongPercent = Round((SongPositionTotalSec/SongLengthTotalSec)*100)
 
   ' Trap Mismatch Errors
   If Err.Number <> 0 Then
      Status = "Error: " & Err.Description
	  Err.Clear
      SetPropertyValue "Music Player - Misc.Status", Status
   Else
      SetPropertyValue "Music Player - Misc.Song Percent", SongPercent
   End If
   Status = ""
   sleep SleepVar
Loop
