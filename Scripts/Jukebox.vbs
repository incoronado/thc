Option Explicit
'On Error Resume Next
Dim HousebotLocation, DOSHousebotLocation, OrderStr, objDB, su, Action, SleepVar

OrderStr = "title"
DOSHousebotLocation = "C:\Progra~2\Housebot\"
HousebotLocation = "C:\Program Files (x86)\housebot\"

Set objDB  = CreateObject("newObjects.sqlite3.dbutf8")
objDB.Open(HousebotLocation & "config\scripts\songs.mlf")
objDB.BusyTimeout=1000
objDB.AutoType = True
objDB.TypeInfoLevel = 4
Set su = CreateObject("newObjects.utilctls.StringUtilities")
PopulateGenreList
objDB.Close
Set objDB  = Nothing
Set su  = Nothing
SetpropertyValue "Jukebox.Jukebox - Import Flag", "Stopped"  
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))

Do
	Sleep SleepVar
	Action = GetPropertyValue ("Jukebox.Action")
	If Action <> "Idle" Then
		MessageHandler(Action)
		SetpropertyValue "Jukebox.Action", "Idle"
	End If
Loop

Set objDB = Nothing

'test Comment

Sub MessageHandler(message)
  Dim a, b, c, SelectedJukebox, SongQueueNo
	
  Set objDB  = CreateObject("newObjects.sqlite3.dbutf8")
  objDB.Open(HousebotLocation & "config\scripts\songs.mlf")
  
  objDB.BusyTimeout=1000
  objDB.AutoType = True
  objDB.TypeInfoLevel = 4
  Set su = CreateObject("newObjects.utilctls.StringUtilities")
  a=split(message,".")
 
  
  Select Case a(0)
	Case "Jukebox"
		b=split(a(3),":")
		SelectedJukebox = GetPropertyValue(GetRemoteNumber(a(1)) & ".Selected Jukebox")
		Select Case b(0)
			'Jukebox.GalaxyTabA1.10.PopulateJukeboxPlaylist
			Case "PopulateJukeboxPlaylist"
			    SongQueueNo = GetPropertyValue("Jukebox" & SelectedJukebox & ".Now Playing - Queue Number")
				PopulatePlaylist SongQueueNo, SelectedJukebox
			Case "PlaySelectedSong"
				PlaySelectedSong b(1)
			Case "LoadSelectedPlaylist"
				LoadSelectedPlaylist(SelectedJukebox)
			Case "PlayNextSong"
				PlayNextSong b(1)	
				'Jukebox.GalaxyTabA1.10.PlayNextSong:1
			Case "PlayPreviousSong"
				PlayPreviousSong b(1)
				'Jukebox.GalaxyTabA1.10.PlayPreviousSong:1
			Case "ToggleRepeat"
				ToggleRepeat b(1)
				'Jukebox.GalaxyTabA1.10.ToggleRepeat:1
			Case "TransferWorkingPlaylist"
			    'Jukebox.GalaxyTabA1.10.TransferWorkingPlaylist:1	
				TransferWorkingPlaylist b(1)
			Case "PopulateLibrarySearchList"
				'Jukebox.GalaxyTabA1.10.PopulateLibrarySearchList
				c=split(GetRemoteNumber(a(1)),"-")
				PopulateLibrarySearchList "False", c(1), GetRemoteNumber(a(1))
			Case "PopulateLibrarySearchListWithDrilldown"
				'Jukebox.GalaxyTabA1.10.PopulateLibrarySearchListWithDrilldown
				c=split(GetRemoteNumber(a(1)),"-")
				PopulateLibrarySearchList "True", c(1), GetRemoteNumber(a(1))
			Case "AddAllSongs2LibrarySongList"
			    'Jukebox.GalaxyTabA1.10.AddAllSongs2LibrarySongList
				AddAllSongs2LibrarySongList GetRemoteNumber(a(1))
			Case "DeleteSelectedSongFromLibrarySongList"
				'Jukebox.GalaxyTabA1.10.DeleteSelectedSongFromLibrarySongList
				DeleteLibrarySongListItem GetRemoteNumber(a(1))
			Case "ClearLibrarySongList"
				'Jukebox.GalaxyTabA1.10.ClearLibrarySongList
				ClearLibrarySongList GetRemoteNumber(a(1))
			Case "ShowLibrarySelectedSong"
				'Jukebox.GalaxyTabA1.10.ShowLibrarySelectedSong
				ShowLibrarySelectedSong GetRemoteNumber(a(1))
			Case "PopulateGenreList"
				PopulateGenreList
				'Jukebox.GalaxyTabA1.10.PopulateGenreList
			Case "Song Select"
				PopulateNameList "True", 0
			Case "Populate Scratchlist"
				PopulateNameList "True", 0
			Case "ToggleEQ"
				'Jukebox.GalaxyTabA1.10.ToggleEQ:1
				ToggleEQ b(1)
			Case "LoadSelectedPlaylistToRemote"
				'Jukebox.GalaxyTabA1.10LoadSelectedPlaylistToRemote
				LoadSelectedPlaylistToRemote GetRemoteNumber(a(1))
			Case "ClearPlayList"
				'Jukebox.GalaxyTabA1.10.ClearPlayList:1
				ClearPlaylist B(1)

		End Select
	
    Case "Song Select"
          PopulateNameList "True", 0
	Case "Populate Scratchlist"
		PopulateNameList "True", 0
	Case "Shuffle Scratchlist"
		ShufflePlaylist(0)
	Case "Shuffle Playlist"
		ShufflePlaylist(SelectedJukebox)
    Case "Add All Playlist"				
		AddAllSongs2PlayList(SelectedJukebox)
	Case "Transfer Working Playlist 1"	
		TransferWorkingPlaylist(1)	
	Case "Transfer Working Playlist 2"	
		TransferWorkingPlaylist(2)	
	Case "Add All Playlist 1"			
		AddAllSongs2PlayList(1)
	Case "Add All Playlist 2"			
		AddAllSongs2PlayList(2)	
	Case "Add All Scratchlist"				
		AddAllSongs2PlayList(0)	
    Case "Add Song Playlist"
		AddSong2PlayList(SelectedJukebox)
	Case "Add Song Scratchlist"
		AddSong2PlayList(0)	
    Case "Clear Playlist"
		ClearPlaylist(SelectedJukebox)
	Case "Clear Scatchlist"
		ClearPlaylist(0)	
	Case "Toggle Repeat"
		ToggleRepeat(SelectedJukebox)
	Case "Clear Scratchlist"	
		ClearPlaylist(0)	
    Case "Clear Playlist 1"	
		ClearPlaylist(1)
	Case "Clear Playlist 2"
		ClearPlaylist(2)
    Case "Create Playlist"
		CreatePlaylist(a(1)) 
    Case "Delete Saved Playlist"
		DeleteSavedPlaylist
	Case "Save Selected Scratchlist"
		SavedSelectedPlaylist(0)
    Case "Save Selected Playlist"
		SavedSelectedPlaylist(SelectedJukebox)
	Case "Load Selected Scratchlist"
		LoadSelectedPlaylist(0)	
    Case "Load Selected Playlist"
		LoadSelectedPlaylist(SelectedJukebox)
    Case "Play Selected Song"
		PlaySelectedSong SelectedJukebox
    Case "Play Next Song"
		PlayNextSong(SelectedJukebox)
	Case "Play Next Song 1"
		PlayNextSong(1)
	Case "ToggleEQ"
		ToggleEQ SelectedJukebox
	Case "Play Next Song 2"
		PlayNextSong(2)
	Case "Play Previous Song"
		PlayPreviousSong(SelectedJukebox)
	Case "Play Previous Song 1"
		PlayPreviousSong(1)
	Case "Play Previous Song 2"	
        PlayPreviousSong(2)
	Case "Import Music Library"
		ImportMusicLibrary
    Case "Song Select No Drilldown"
		PopulateNameList "False", 0
    Case "Populate Genre List"
		PopulateGenreList
		
	Case "Populate Playlist"
		PopulatePlaylist 0, SelectedJukebox
	Case "Populate Playlist 1"
		SetPropertyValue "Jukebox.Debug", "Got Here"
		PopulatePlaylist 0, 1
	Case "Populate Playlist 2"
		PopulatePlaylist 0, 2
		
    Case "Populate Saved Playlist"
		PopulateSavedPlayLists
	Case "Selected Song Detail"
		ShowLibrarySelectedSong
	Case "Delete Playlist Item"
		DeletePlaylistItem(SelectedJukebox)
	Case "Delete Playlist 1 Item"
		DeletePlaylistItem(1)
	Case "Delete Playlist 2 Item"
		DeletePlaylistItem(2)
	Case "Delete Scratchlist Item"
		DeletePlaylistItem(0)	
    Case "Song Percent"
		SongPercent(SelectedJukebox)
	Case "OpenKeyboard"
		OpenKeyboard
	Case "OpenGenreList"
		OpenGenreList
	Case "CloseGenrePanel"
		CloseGenrePanel
  End Select 
  objDB.Close
  Set objDB = Nothing
  Set su = nothing
End Sub


Sub ToggleRepeat(JukeboxNo)
    If GetPropertyValue ("Jukebox" & CStr(JukeboxNo) & ".Jukebox - Repeat") = "On" Then
	   SetPropertyValue "Jukebox" & CStr(JukeboxNo) & ".Jukebox - Repeat", "Off"
	Else
	   SetPropertyValue "Jukebox" & CStr(JukeboxNo) & ".Jukebox - Repeat", "On"
	End If
End Sub




Function OpenGenreList
	PopulateGenreList
	OpenRemotePanel("Genre")
End Function

Function CloseGenrePanel
	CloseRemotePanel("Genre")
	'SetpropertyValue "Jukebox.Action", "Song Select No Drilldown"
End Function

Sub OpenKeyboard
	OpenRemotePanel("Keyboard") 
End Sub


Sub NowPlayingThumbnail(ListNo)
   Dim CmdStr, JavaStr, CurrentSong, ImageFile, objFSO, oShell, ImageFileNoQuotes , error_code, i, return_code
   ' The following sleep is important to make sure the dust has seleted on  "Music Player N.Play Source"
   sleep 100
   CurrentSong = GetPropertyValue("Music Player " & CStr(ListNo) & ".Play Source")
   ImageFile = "config\themes\THC\Current Song" & ListNo & ".jpg"
   ImageFileNoQuotes = ImageFile
   
   ' Copy Default image so if embedded image not found displays default image.
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   'objFSO.CopyFile HousebotLocation & "config\themes\THC\cdi.jpg" , HousebotLocation & ImageFile, True
      
   CmdStr = "ffmpeg -loglevel quiet -y -i """ & CurrentSong & """ -vf scale=530:530 """ & ImageFile & ""
   
   Set oShell = CreateObject("WScript.Shell")
   error_code = oShell.Run("cmd /c " & CmdStr, 0, true)
   sleep 150
   return_code = SetPropertyValue("Jukebox" & CStr(ListNo) & ".Now Playing Image", ImageFileNoQuotes)
   'SetPropertyValue "Jukebox.Debug", return_code 
   sleep 150
   return_code = SetPropertyValue("Jukebox" & CStr(ListNo) & ".Now Playing Image", ImageFileNoQuotes)
   
   
   ' Close & Free Objects
   Set objFSO = Nothing
   Set oShell = Nothing
  
End Sub

Sub SelectedSongThumbnail (Remote)
   Dim CmdStr, JavaStr, CurrentSong, ImageFile, objFSO, return_code, error_code, a, RemoteNumber, ImageFileNoQuotes
   'sleep 50
   CurrentSong = GetPropertyValue(Remote & ".Library - File")
   ImageFile = GetPropertyValue(Remote & ".Library - Selected Image")
   ImageFileNoQuotes = ImageFile
   
   ' Copy Default image so if embedded image not found displays default image.
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   objFSO.CopyFile HousebotLocation & "config\themes\THC\cdi-sm.png" , HousebotLocation & ImageFile, True
    CmdStr = "ffmpeg -loglevel quiet -y -i """ & CurrentSong & """ -vf scale=75:75 """ & ImageFile & ""
	Dim oShell 
	Set oShell = CreateObject( "WScript.Shell" )
	error_code = oShell.Run("cmd /c " & CmdStr, 0, true)
	sleep 50 
	return_code = SetPropertyValue(Remote & ".Library - Selected Image", ImageFileNoQuotes)
	sleep 50 
	return_code = SetPropertyValue(Remote & ".Library - Selected Image", ImageFileNoQuotes)
	' Close & Free Objects
   Set objFSO = Nothing
   Set oShell = Nothing

End Sub



Sub PlayNextSong(ListNo)
	Dim queuenmbr, SqlStr, r
   'msgbox "Got Here"
   	If Trim(GetPropertyValue ("Jukebox" & CStr(ListNo) & ".Jukebox - Playing Song")) <> "" Then
		queuenmbr = GetPropertyValue("Jukebox" & CStr(ListNo) & ".Now Playing - Queue Number") 
	 	SqlStr = "select * from songqueue where playlistid = " & Cstr(ListNo)
		Set r = objDB.Execute(SqlStr)
		SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Queue Total", r.Count
		'MsgBox r.Count & ", " & Cint(queuenmbr)  
			If r.Count > Cint(queuenmbr) Then
				SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Queue Number", queuenmbr + 1
				PopulatePlayList queuenmbr + 1 ,ListNo
				sleep 50
				'SetPropertyValue "Jukebox.Debug", "List Number->" & ListNo & "Playlist^" & queuenmbr + 1 & "^" & PlayListSongID(queuenmbr + 1,ListNo)		 
				SetPropertyValue "Jukebox" & CStr(ListNo) & ".Jukebox - Playing Song", "Playlist^" & queuenmbr + 1 & "^" & PlayListSongID(queuenmbr + 1,ListNo)
				sleep 150
				PlaySelectedSong(Listno)
			Else
				If GetPropertyValue ("Jukebox" & CStr(ListNo) & ".Jukebox - Repeat") = "On" Then
					SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Queue Number", 1
					PopulatePlayList 1, ListNo
					sleep 50
					SetPropertyValue "Jukebox" & CStr(ListNo) & ".Jukebox - Playing Song", "Playlist^1" & "^" & PlayListSongID(1,ListNo)
					'SetPropertyValue "Jukebox.Jukebox - Playing Song", "Playlist^" & queuenmbr + 1 & "^" & PlayListSongID(queuenmbr + 1,ListNo)
					sleep 150
					PlaySelectedSong(ListNo)
				End If
			End If
			Set r = Nothing
	End If
   'sleep 300
	NowPlayingThumbnail(ListNo)
End Sub

Sub PlayPreviousSong(ListNo)
	Dim queuenmbr, SqlStr, r
	If Trim(GetPropertyValue ("Jukebox" & CStr(ListNo) & ".Jukebox - Playing Song")) <> "" Then
  	  queuenmbr = GetPropertyValue("Jukebox" & CStr(ListNo) & ".Now Playing - Queue Number")
	  
	  ' Right(GetPropertyValue ("Jukebox.Jukebox - Playing Song"),Len(GetPropertyValue ("Jukebox.Jukebox - Playing Song")) - 9)
      SqlStr = "select * from songqueue where playlistid = " & CStr(ListNo)
      Set r = objDB.Execute(SqlStr)
      SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Queue Total", r.Count
      If Cint(queuenmbr) > 1  Then
         SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Queue Number", queuenmbr - 1
         'MsgBox "Got Milk"
         PopulatePlayList queuenmbr - 1, ListNo
		 sleep 50		 
         SetPropertyValue "Jukebox" & CStr(ListNo) & ".Jukebox - Playing Song", "Playlist^" & queuenmbr - 1 & "^" & PlayListSongID(queuenmbr - 1,ListNo)
		 sleep 150
		 PlaySelectedSong(ListNo)
      End If 
      Set r = Nothing
   End If
   'sleep 300
   'NowPlayingThumbnail(ListNo)
End Sub

Sub PlaySelectedSong(ListNo)
  
  Dim data, SqlStr, r
  If Trim(GetPropertyValue ("Jukebox" & CStr(ListNo) & ".Jukebox - Playing Song")) <> "" Then
     data = Split(GetPropertyValue ("Jukebox" & CStr(ListNo) & ".Jukebox - Playing Song"),"^")
	 SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Queue Number", data(1)
	 'SqlStr = "select * from songs where id = (select songid from songqueue where id = " & data(1) & ")"
	 SqlStr = "select * from songs where id = " & data(2) 
     'MsgBox SqlStr
     Set r = objDB.Execute(SqlStr)
     SetPropertyValue "Music Player " & CStr(ListNo) & ".Play Source", r(1)("file")
	 SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Genre", r(1)("genre")
	 SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Album", r(1)("album")
	 SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Artist", r(1)("artist")
	 SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Title", r(1)("title")
	 SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Track", r(1)("track")
	 SetPropertyValue "Jukebox" & CStr(ListNo) & ".Now Playing - Year", r(1)("year")
	 'SetPropertyValue "Music Player " & CStr(ListNo) & ".Status", "Playing"
	 Set r = Nothing
  End If
  
  NowPlayingThumbnail(ListNo)
  
  'sleep 500
End Sub

Sub ShowLibrarySelectedSong(Remote)
	Dim a, RemoteNumber, SqlStr, r
	
	a=split(Remote,"-")
	RemoteNumber=a(1)
	
    If (GetPropertyValue (Remote & ".Jukebox - Tag Type") = "title") And (Trim(GetPropertyValue(Remote & ".Jukebox - Selected Name"))) <> "" Then
		'sleep 150
		 a=split(GetPropertyValue(Remote & ".Jukebox - Selected Name"),"^")
		 SqlStr = "select * from songs where id = " & a(1)
		 'MsgBox SqlStr
		 Set r = objDB.Execute(SqlStr)
		 'SetPropertyValue "Music Player.Play Source", r(1)("file")
		 SetPropertyValue Remote & ".Library - Genre", r(1)("genre")
		 SetPropertyValue Remote & ".Library - Album", r(1)("album")
		 SetPropertyValue Remote & ".Library - Artist", r(1)("artist")
		 SetPropertyValue Remote & ".Library - Title", r(1)("title")
		 SetPropertyValue Remote & ".Library - Track", r(1)("track")
		 SetPropertyValue Remote & ".Library - Year", r(1)("year")
		 SetPropertyValue Remote & ".Library - File", r(1)("file")
         Set r = Nothing		 
		 sleep 5
		 SelectedSongThumbnail Remote
		 'sleep 50
	 ElseIf (GetPropertyValue (Remote & ".Jukebox - Tag Type") = "album") And (Trim(GetPropertyValue(Remote & ".Jukebox - Selected Name"))) <> "" Then	 
	 
	 
	 End If
End Sub

Function ProcessSongSelect
  If GetPropertyValue ("Jukebox.Jukebox - Tag Type") = "album" Then
    'MsgBox "Album"
  ElseIf GetPropertyValue ("Jukebox.Jukebox - Tag Type") = "artist" Then
    'MsgBox "Artist"
  End if
End Function


Function GetPropertyValueTrapErrors(PropertyName, ErrorProperty)
On Error Resume Next
Dim Status, i, RetryCount
RetryCount = 5
   For i=1 To RetryCount
      Status = ""
	  TmpPropertyValue = GetPropertyValue(PropertyName)
	  If Err.Number <> 0 Then
         Status = PropertyName & " Read Error: " & Err.Description
	     Err.Clear
		 If i = 5 Then
            SetPropertyValue ErrorProperty, Status
            'Return Error
			TmpPropertyValue = "Error"
		 End if
	  Else
	     ' Good Read
         Exit For
      End If
   Next
   'Return 
   GetPropertyValueTrapErrors = TmpPropertyValue
End Function




Sub SavedSelectedPlaylist(ListNo)
  Dim SelectedPlaylist, SqlStr, r
  ' delete existing songs from selected savedplaylist 
  SelectedPlaylist = GetPropertyValueTrapErrors ("Jukebox.Jukebox - Selected Playlist", "Jukebox.Error")
  If SelectedPlaylist <> "Error" Then 
     objDB.Execute("delete from savedplaylists where playlistid  = " & Right(SelectedPlaylist,Len(SelectedPlaylist) - 14))
  End If

  Set r = objDB.Execute("select * from songqueue where playlistid  = " & (ListNo))
  For Row = 1 To r.Count
	 SelectedPlaylist = GetPropertyValueTrapErrors ("Jukebox.Jukebox - Selected Playlist", "Jukebox.Error")
	 If SelectedPlaylist <> "Error" Then 
        objDB.Execute su.Sprintf("INSERT INTO savedplaylists (playlistid, songid) VALUES (%Nq, %Nq)", Right(SelectedPlaylist,Len(SelectedPlaylist) - 14), r(Row)("songid") )
	 End if
  Next
  Set r = Nothing
  Sleep 500
  PopulateSavedPlayLists
End Sub

Sub ToggleEQ(JukeboxNo)
	If GetPropertyValue("Music Player " & JukeboxNo & ".Equalizer Enabled") = 1 Then
		SetPropertyValue "Music Player " & JukeboxNo & ".Equalizer Enabled", 0
	Else
		SetPropertyValue "Music Player " & JukeboxNo & ".Equalizer Enabled", 1
	End if
End Sub

Sub LoadSelectedPlaylistToRemote(Remote)
  Dim SqlStr, r, Row
  ' delete existing songs from selected savedplaylist 
  SqlStr = "delete from songqueue where playlistid = 0"
  objDB.Execute(SqlStr)
   SetPropertyValue "Jukebox.Debug", "Remote-" & CStr(Remote) & ".Jukebox - Selected Playlist"

  SqlStr = "select * from savedplaylists where playlistid  = " & Right(GetPropertyValue(Remote & ".Jukebox - Selected Playlist"),Len(GetPropertyValue(Remote & ".Jukebox - Selected Playlist")) - 14)
  Set r = objDB.Execute(SqlStr)
  For Row = 1 To r.Count 
     objDB.Execute su.Sprintf("INSERT INTO songqueue (playlistid, songid) VALUES (%Nq, %Nq)", 0, r(Row)("songid") )
  Next
  Set r = Nothing
  sleep 500
  PopulateLibrarySongList 0, Remote
End Sub











Sub LoadSelectedPlaylist(ListNo)
  Dim SqlStr, r, Row
  ' delete existing songs from selected savedplaylist 
  SqlStr = "delete from songqueue where playlistid = " & ListNo 
  objDB.Execute(SqlStr)

  SqlStr = "select * from savedplaylists where playlistid  = " & Right(GetPropertyValue ("Jukebox.Jukebox - Selected Playlist"),Len(GetPropertyValue ("Jukebox.Jukebox - Selected Playlist")) - 14)
  Set r = objDB.Execute(SqlStr)
  For Row = 1 To r.Count 
     objDB.Execute su.Sprintf("INSERT INTO songqueue (playlistid, songid) VALUES (%Nq, %Nq)", ListNo, r(Row)("songid") )
  Next
  Set r = Nothing
  sleep 500
  PopulatePlayList 0, ListNo
End Sub


Sub DeleteSavedPlaylist
  SqlStr = "delete from playlistnames where id = " & Right(GetPropertyValue ("Jukebox.Jukebox - Selected Playlist"),Len(GetPropertyValue ("Jukebox.Jukebox - Selected Playlist")) - 14)
  objDB.Execute(SqlStr)
  SqlStr = "delete from savedplaylists where playlistid  = " & Right(GetPropertyValue ("Jukebox.Jukebox - Selected Playlist"),Len(GetPropertyValue ("Jukebox.Jukebox - Selected Playlist")) - 14)
  objDB.Execute(SqlStr)
 'MusicList = Row & vbCR
  Sleep 500
  PopulateSavedPlayLists
End Sub


Sub CreatePlaylist (Remote)
   SqlStr = su.Sprintf("select * from playlistnames where playlistname = %Nq", GetPropertyValue (Remote & " System.Jukebox - New Playlist"))
   'msgbox SqlStr
   Set r = objDB.Execute(SqlStr)

   If (r.count = 0) Then
      objDB.Execute su.Sprintf("INSERT INTO playlistnames (playlistname) VALUES (%Nq)", GetPropertyValue (Remote & " System.Jukebox - New Playlist"))
   Else
       objDB.Execute su.Sprintf("INSERT INTO playlistnames (playlistname) VALUES (%Nq)", GetPropertyValue (Remote & " System.Jukebox - New Playlist") & "1")
   End If

   Set r = Nothing

   ' Clear Kbd Buffer
   SetPropertyValue Remote & " System.Jukebox - New Playlist", ""
   SetPropertyValue Remote & " System.Jukebox - New Playlist", "^"
   sleep 500
   PopulateSavedPlayLists
End Sub

Sub PopulateSavedPlayLists
	Dim PlayList, SqlStr, r, Row, s
	PlayList = ""
	SqlStr = "select * from playlistnames"
	Set r = objDB.Execute(SqlStr)
	If r.Count = 0 Then
	   PlayList     = " " & vbTAB & "No Playlists Saved" & vbLF
	End if 
	For Row = 1 To r.count
	  SqlStr = "select * from savedplaylists where playlistid = " & r(Row)("id")
	  Set s = objDB.Execute(SqlStr)
	  PlayList = PlayList & "SavedPlaylist^" & r(Row)("id") & vbTAB & r(Row)("playlistname") & vbTAB & " (" & s.Count & ")" & vbLF
	 'MusicList = Row & vbCR
	Next
	Set r = Nothing
	SetPropertyValue "Jukebox.Jukebox - Saved Playlists", Left(PlayList,Len(PlayList)-1)
End Sub

Sub PopulateLibrarySearchList(drilldown,ListNo,Remote)
	Dim MusicList, tagType, SortLetter, usewhere, r, WhereStr, GenreStr, Row
	Dim SearchLetterStr, SearchStr, SqlStr, OrderStr
	'SetModeState "Block Library Selected Song", "Inactive"
	tagType = GetPropertyValue (Remote & ".Jukebox - Tag Type")
	usewhere = 0
	GenreStr = ""
	SearchLetterStr = ""
	SearchStr = ""
	MusicList = ""
	'SetPropertyValue "JukSetModeStateboxebox.Selected Name", ""
	
	If drilldown = "True"  Then
		If tagType= "album" Then
			SqlStr = "select * from songs where album = '" & Replace(Mid(GetPropertyValue (Remote & ".Jukebox - Double-Click Name"),11,Len(GetPropertyValue (Remote & ".Jukebox - Selected Name")) -10),"'","''") & "' order by CAST(track AS INTEGER)"
	    	'MsgBox SqlStr
			SetModeState "Block PopulateLibrarySearchList", "Inactive"
			sleep 5
			SetPropertyValue Remote & ".Jukebox - Tag Type", "title"
			
			SetModeState "Block PopulateLibrarySearchList", "Active"
			OrderStr = "album"
			tagType = "title"
		ElseIf tagType= "artist" Then
			SqlStr = "select * from songs where artist = '" & Replace(Mid(GetPropertyValue (Remote & ".Jukebox - Double-Click Name"),11,Len(GetPropertyValue (Remote & ".Jukebox - Selected Name")) -10), "'","''") & "' order by title collate nocase"
			'MsgBox "2"
			SetPropertyValue "Remote-1.Debug", SqlStr
			SetModeState "Block PopulateLibrarySearchList", "Inactive"
			sleep 5
			SetPropertyValue Remote & ".Jukebox - Tag Type", "title"
			SetModeState "Block PopulateLibrarySearchList", "Active"
			OrderStr = "artist"
			tagType = "title"
		Else
		    SetModeState "Block PopulateLibrarySearchList", "Inactive"
			AddSong2LibrarySongList ListNo, Remote 
			SetModeState "Block PopulateLibrarySearchList", "Active"
			Exit Sub
		End If
	Else
		If GetPropertyValue (Remote & ".Jukebox - Selected Genre") <> "" Then
			usewhere = usewhere + 1
			If usewhere > 1 Then
				GenreStr = " AND "
			End if
			GenreStr = GenreStr & "genre = '" & GetPropertyValue (Remote & ".Jukebox - Selected Genre") & "'"
			sleep 5
			PopulateGenreList
		End If
  
		If GetPropertyValue (Remote & ".Jukebox - Search") <> "*" Then
			usewhere = usewhere + 1
			If usewhere > 1 Then
				SearchLetterStr = " AND "
			End if			
			SearchLetterStr = SearchLetterStr & tagType & " like '" & GetPropertyValue (Remote & ".Jukebox - Search") & "%'"
		End if
   
		If usewhere >= 1 Then
			WhereStr = " where "
		Else 
			WhereStr = ""
		End If
		
		If tagType = "title" Then
			SqlStr = "select * from songs" & WhereStr & GenreStr & SearchLetterStr & " order by title" 
			' & OrderStr & " collate nocase"
			'MsgBox "3: " & SqlStr
		Else
			SqlStr = "select distinct " & tagType & " from songs" & WhereStr & GenreStr & SearchLetterStr & " order by " & tagType & " collate nocase"
			'MsgBox "4: " & SqlStr	
		End If
	End if
	SetPropertyValue Remote & ".Jukebox - SQL Search", SqlStr
	'MsgBox SqlStr
	
	'SetPropertyValue Remote & ".Library - Filtered Song List", "*S-" & vbTAB & "Building List..."
	Set r = objDB.Execute(SqlStr)
	'MsgBox StartCount & ", " & EndCount & "," & r.Count
	
	If r.Count = 0 Then
		MusicList     = vbTAB & "No Records Found" & vbLF
	Else 
		For Row = 1 To r.Count
			If tagType = "title" Then
				MusicList = MusicList & "MusicList^" & r(Row)("id") & vbTAB & r(Row)(tagType) & vbLF
			Else
				MusicList = MusicList & "MusicList^" & r(Row)(tagType) & vbTAB & r(Row)(tagType) & vbLF
			End if
		Next
	End if
	Set r = Nothing
	'MsgBox Len(MusicList)-1
	'SetModeState "Block Jukebox Search Filter", "Inactive"
	SetPropertyValue Remote & ".Library - Filtered Song List", Left(MusicList,Len(MusicList)-1) 
	'SetModeState "Block Jukebox Search Filter", "Active"
	
End Sub




Sub PopulateNameList(drilldown,ListNo,Remote)
	Dim MusicList, tagType, SortLetter, usewhere, r, WhereStr, GenreStr
	Dim SearchLetterStr, SearchStr, SqlStr, OrderStr
	
	tagType = GetPropertyValue (Remote & ".Jukebox - Tag Type")
	usewhere = 0
	GenreStr = ""
	SearchLetterStr = ""
	SearchStr = ""
	MusicList = ""
	'SetPropertyValue "Jukebox.Selected Name", ""
	If drilldown = "True"  Then
		If tagType= "album" Then
			SqlStr = "select * from songs where album = '" & Replace(Mid(GetPropertyValue (Remote & " System.Jukebox - Double-Click Name"),11,Len(GetPropertyValue (Remote & " System.Jukebox - Selected Name")) -10),"'","''") & "' order by CAST(track AS INTEGER)"
	    	'MsgBox SqlStr
			SetModeState "Allow Run Jukebox Script", "Inactive"
			sleep 5
			SetPropertyValue Remote & ".Jukebox - Tag Type", "title"
			SetModeState "Allow Run Jukebox Script", "Active"
			OrderStr = "album"
			tagType = "title"
		ElseIf tagType= "artist" Then
			SqlStr = "select * from songs where artist = '" & Replace(Mid(GetPropertyValue (Remote & " System.Jukebox - Double-Click Name"),11,Len(GetPropertyValue (Remote & " System.Jukebox - Selected Name")) -10), "'","''") & "' order by title collate nocase"
			'MsgBox "2"
			SetModeState "Allow Run Jukebox Script", "Inactive"
			sleep 5
			SetPropertyValue Remote & " System.Jukebox - Tag Type", "title"
			SetModeState "Allow Run Jukebox Script", "Active"
			OrderStr = "artist"
			tagType = "title"
		Else
			AddSong2PlayList ListNo,Remote 
			Exit Sub
		End If
	Else
		If GetPropertyValue (Remote & " System.Jukebox - Selected Genre") <> "" Then
			usewhere = usewhere + 1
			If usewhere > 1 Then
				GenreStr = " AND "
			End if
			GenreStr = GenreStr & "genre = '" & GetPropertyValue (Remote & " System.Jukebox - Selected Genre") & "'"
			sleep 5
			PopulateGenreList
		End If
  
		If GetPropertyValue (Remote & ".Jukebox - Search") <> "*" Then
			usewhere = usewhere + 1
			If usewhere > 1 Then
				SearchLetterStr = " AND "
			End if			
			SearchLetterStr = SearchLetterStr & tagType & " like '" & GetPropertyValue (Remote & ".Jukebox - Search") & "%'"
		End if
   
		If usewhere >= 1 Then
			WhereStr = " where "
		Else 
			WhereStr = ""
		End If
		If tagType = "title" Then
			SqlStr = "select * from songs" & WhereStr & GenreStr & SearchLetterStr & " order by title" 
			' & OrderStr & " collate nocase"
			'MsgBox "3: " & SqlStr
		Else
			SqlStr = "select distinct " & tagType & " from songs" & WhereStr & GenreStr & SearchLetterStr & " order by " & tagType & " collate nocase"
			'MsgBox "4: " & SqlStr	
		End If
	End if
	SetPropertyValue Remote & " System.Jukebox - Last SQL", SqlStr
	'MsgBox SqlStr
	Set r = objDB.Execute(SqlStr)
	'MsgBox StartCount & ", " & EndCount & "," & r.Count
	If r.Count = 0 Then
		MusicList     = vbTAB & "No Records Found" & vbLF
	Else 
		For Row = 1 To r.Count
			If tagType = "title" Then
				MusicList = MusicList & "MusicList^" & r(Row)("id") & vbTAB & r(Row)(tagType) & vbLF
			Else
				MusicList = MusicList & "MusicList^" & r(Row)(tagType) & vbTAB & r(Row)(tagType) & vbLF
			End if
		Next
	End if
	Set r = Nothing
	'MsgBox Len(MusicList)-1
	SetPropertyValue Remote & ".Libray - Song List", Left(MusicList,Len(MusicList)-1) 
	
End Sub


Sub DeletePlaylistItem(ListNo)
    Dim data
    If ListNo = 0 Then
   		data = Split(GetPropertyValue ("GalaxyProTab1 System.Jukebox - Playing Song"),"^")
	Else
		data = Split(GetPropertyValue ("Jukebox" & CStr(ListNo) & ".Jukebox - Playing Song"),"^")
	End if	

   objDB.Execute("delete from songqueue where id = " & data(3)) 
   sleep 250
   PopulatePlayList 0, ListNo
End Sub


Sub DeleteLibrarySongListItem(Remote)
    Dim data
    data = Split(GetPropertyValue (Remote & ".Library - Song List Selected Song"),"^")
	objDB.Execute("delete from librarysonglist where id = " & data(3)) 
	sleep 250
	PopulateLibrarySongList 0, Remote 
End Sub


Sub AddSong2PlayList(ListNo,Remote)
	 objDB.Execute su.Sprintf("INSERT INTO songqueue (playlistid, songid) VALUES (%Nq,%Nq)",CInt(ListNo),Right(GetPropertyValue ("Jukebox.Jukebox - Double-Click Name"),Len(GetPropertyValue ("Jukebox.Jukebox - Double-Click Name")) - 10))
	 sleep 50
    PopulatePlayList 0, ListNo
End Sub

Sub AddSong2LibrarySongList(ListNo,Remote)
	 objDB.Execute su.Sprintf("INSERT INTO librarysonglist (playlistid, songid) VALUES (%Nq,%Nq)",CInt(ListNo),Right(GetPropertyValue (Remote & ".Jukebox - Double-Click Name"),Len(GetPropertyValue (Remote & ".Jukebox - Double-Click Name")) - 10))
	 sleep 50
	 PopulateLibrarySongList 0, Remote
End Sub

Sub AddAllSongs2LibrarySongList(Remote)
	Dim a, RemoteNumber, SqlStr, r, Row
	SetPropertyValue Remote & ".Library - Song List", "*S-" & vbTAB & "Building List..."
	' Extract Remote Number from Remote Name
	a=split(Remote,"-")
	RemoteNumber=a(1)
	
	If GetPropertyValue(Remote & ".Jukebox - Tag Type") = "title" Then
		SqlStr = GetPropertyValue(Remote & ".Jukebox - SQL Search")
		Set r = objDB.Execute(SqlStr)
		For Row = 1 To r.Count
			objDB.Execute su.Sprintf("INSERT INTO librarysonglist (playlistid, songid) VALUES (%Nq,%Nq)",CInt(RemoteNumber),r(Row)("id"))
		Next
		Set r = Nothing
	End If

	sleep 50
    PopulateLibrarySongList 0, Remote
End Sub


Sub AddAllSongs2PlayList(ListNo)
	
   If GetPropertyValue("Jukebox.Jukebox - Tag Type") = "title" Then
      SqlStr = GetPropertyValue("Jukebox.Jukebox - Last SQL")
      Set r = objDB.Execute(SqlStr)
      For Row = 1 To r.Count
         objDB.Execute su.Sprintf("INSERT INTO songqueue (playlistid, songid) VALUES (%Nq,%Nq)",CInt(ListNo),r(Row)("id"))
      Next
	  Set r = Nothing
	End If

	sleep 50
    PopulatePlayList 0, ListNo
End Sub



Function PopulateGenreList
	Dim GenreList, Row, SqlStr, r, SelectedPrefix
	
	SqlStr = "select * from genre where genre <> '' and genre <> '<empty>' order by genre"
	Set r = objDB.Execute(SqlStr)
	'MsgBox SqlStr

	If r.Count = 0 Then
	   GenreList     = "*S-" & "All" & vbTAB & "No Records Found"
	End if 

	For Row = 1 To r.Count
		If (GetPropertyValue ("Jukebox.Selected Genre") = r(Row)("genre")) Then
			SelectedPrefix =  "*S-"
		Else
			SelectedPrefix =  ""
		End if
		If Row = 1 Then
			GenreList = SelectedPrefix & "" & vbTAB & "Select All"
		Else
			GenreList = GenreList & vbLF & SelectedPrefix & r(Row)("genre") & vbTAB & r(Row)("genre")
		End if
	Next

	Set r = Nothing
	SetPropertyValue "Jukebox.Jukebox - Genre List", GenreList

End Function

Sub TransferWorkingPlaylist(ListNo)
    SetpropertyValue "Jukebox.Jukebox - Import Messages", "Saving Song List of Jukebox " & CStr(ListNo)
	Dim SqlStr, r, queuenmbr
	queuenmbr = GetPropertyValue("Jukebox" & CStr(ListNo) & ".Now Playing - Queue Number")
	SqlStr = "CREATE TEMPORARY TABLE temp_table AS SELECT * FROM librarysonglist; UPDATE temp_table SET id=NULL, playlistid=" & ListNo & "; INSERT INTO songqueue SELECT * FROM temp_table; DROP TABLE temp_table;"
	Set r = objDB.Execute(SqlStr)
	Set r = Nothing
	
	SetpropertyValue "Jukebox.Jukebox - Import Messages", " Populating Jukebox " & CStr(ListNo) & " Playlist"
	PopulatePlayList queuenmbr,ListNo
	SetpropertyValue "Jukebox.Jukebox - Import Messages", " Finished Transfering Playlist to Jukebox " & CStr(ListNo)
End Sub


Sub PopulateLibrarySongList(selectedtag,Remote)

   ' If selectedtag = 0 process as normal
   Dim PlayList, SqlStr, r, Row, RemoteNumber, a
   
   PlayList = ""
   'Extract Remote Number from Remote Name
   a = split(Remote,"-")
   RemoteNumber = a(1)
   
	SqlStr = "select librarysonglist.id, songs.title, songs.album, songs.track, librarysonglist.songid  from librarysonglist left join songs on librarysonglist.songid = songs.id where librarysonglist.playlistid = " & RemoteNumber & " order by songs.album asc, CAST(songs.track AS int) asc"
	Set r = objDB.Execute(SqlStr)
   
	If r.Count = 0 Then
		PlayList     = " " & vbTAB & "Playlist Empty" & vbLF
	Else
		For Row = 1 To r.count
			If CInt(selectedtag) = Row Then
				PlayList = PlayList & "*S-" & "PlayList^" & Row & "^" & r(Row)("songid") & "^" & r(Row)("id") & vbTAB & r(Row)("title") &vbLF
			Else
				PlayList = PlayList & "PlayList^" & Row & "^" & r(Row)("songid") & "^" & r(Row)("id") & vbTAB & r(Row)("title") &vbLF
			End if
			'MusicList = Row & vbCR
		Next
	End If

	SetPropertyValue Remote & ".Library - Song List", Left(PlayList,Len(PlayList)-1) 
	Set r = Nothing
End Sub





Sub PopulatePlayList(selectedtag,ListNo)
   ' If selectedtag = 0 process as normal
   ' If ListNo = 0 

   Dim PlayList, SqlStr, r, Row
   PlayList = ""
  
   SqlStr = "select songqueue.id, songs.title, songs.album, songs.track, songqueue.songid  from songqueue left join songs on songqueue.songid = songs.id where songqueue.playlistid = " & ListNo & " order by songs.album asc, CAST(songs.track AS int) asc"
   Set r = objDB.Execute(SqlStr)
   
    If r.Count = 0 Then
      PlayList     = " " & vbTAB & "Playlist Empty" & vbLF
	Else
      For Row = 1 To r.count
         If CInt(selectedtag) = Row Then
            PlayList = PlayList & "*S-" & "PlayList^" & Row & "^" & r(Row)("songid") & "^" & r(Row)("id") & vbTAB & r(Row)("title") &vbLF
         Else
            PlayList = PlayList & "PlayList^" & Row & "^" & r(Row)("songid") & "^" & r(Row)("id") & vbTAB & r(Row)("title") &vbLF
         End if
         'MusicList = Row & vbCR
      Next
	End If
    
	'If ListNo = 0 Then
	'	SetPropertyValue "GalaxyProTab1 System.Jukebox - Playlist", Left(PlayList,Len(PlayList)-1) 
	'Else
		SetPropertyValue "Remote-" & CStr(ListNo) & ".Library - Song List", Left(PlayList,Len(PlayList)-1) 
	'End If
   Set r = Nothing
End Sub

Function PlayListSongID(PlayListID,ListNo)
   Dim Playlist, LeftStrPos, RightStr, RightStrPos, PlaylistRecordStr, data 
   PlayList = GetPropertyValue("Jukebox" & CStr(ListNo) & ".Jukebox - Playlist")
   ' MsgBox "PlayList^" & PlayListID & "^"
    'SetPropertyValue "Jukebox.Debug", PlayListID
   LeftStrPos = InStr(PlayList,"PlayList^" & PlayListID & "^")
   RightStr = Mid(PlayList,LeftStrPos)
   RightStrPos = InStr(RightStr,vbTAB)
   PlaylistRecordStr = Left(RightStr,RightStrPos-1)
   data = Split(PlaylistRecordStr,"^") 
   PlayListSongID=data(2)
End Function

Function ClearPlaylist(ListNo)
   objDB.Execute su.Sprintf("delete from songqueue where playlistid = " & CStr(ListNo))
   sleep 150
   PopulatePlayList 0, ListNo
End Function


Sub ClearLibrarySongList(Remote)
	Dim a, RemoteNumber
	a=split(Remote,"-")
	RemoteNumber = a(1)
	objDB.Execute su.Sprintf("delete from librarysonglist where playlistid = " & RemoteNumber)
	sleep 150
	PopulateLibrarySongList 0, Remote 
End Sub


Sub ShufflePlaylist(ListNo)
   Set r = objDB.Execute("select playlistid, songid from songqueue where playlistid = " & Cstr(ListNo) & " order by random()")
   Set ra = objDB.Execute("delete from songqueue where playlistid = " & Cstr(ListNo)) 
   For Row = 1 To r.Count
       objDB.Execute su.Sprintf("INSERT INTO songqueue (playlistid, songid) VALUES (%Nq,%Nq)",ListNo,r(Row)("songid"))
   Next
   Set r = Nothing
   sleep 150
   PopulatePlayList 0, ListNo
End Sub







Sub ImportMusicLibrary

	Dim objFSO , objTextFile, oShell, r, s, objFile 
	Dim CmdStr, RetCode, strNextLine, TitleStr, ArtistStr, AlbumStr, TrackStr, GenreStr, YearStr, FileStr, SqlStr, song_id, str_pos, data, CountStr, Row

	SetpropertyValue "Jukebox.Jukebox - Import Flag", "Running"

	CmdStr = """" & DOSHousebotLocation & "config\scripts\songs.cmd""" 

	'CmdStr = "c:\Progra~1\housebot\config\scripts\id3 -2 -q ""Title  :" & Chr(37) & "t\nArtist : %%a\nAlbum  : %%l\nTrack  : %%n\nGenre  : %%g\nYear   : %%y\nFile   : %%p%%f\nCount  : %%X\n"" -R ""e:\music\*.mp3"""

	 '-- Execute the command and redirect the output to the file
	SetpropertyValue "Jukebox.Jukebox - Import Messages", "Building Tag List - Please Wait"

	Set oShell = CreateObject( "WScript.Shell" )
	oShell.Run CmdStr, 0, True
	sleep 100
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile(HousebotLocation & "config\scripts\songs.tmp", 1)

	'SetPropertyValue "Import Music Library.Execute Program", "Yes"

	'Clear Tables
	' objDB.Execute su.Sprintf("DELETE FROM SONGS")
	objDB.Execute su.Sprintf("DELETE FROM GENRE")
	Do Until objTextFile.AtEndOfStream
		strNextLine = objTextFile.Readline
		' Process song record 
		If Left(strNextLine,9) = "Title  : " Then
		   TitleStr =  Mid(strNextLine,10, len(strNextLine) - 9)
		   strNextLine = objTextFile.Readline
		   ArtistStr =  Mid(strNextLine,10, len(strNextLine) - 9)
		   strNextLine = objTextFile.Readline
		   AlbumStr =  Mid(strNextLine,10, len(strNextLine) - 9)
		   strNextLine = objTextFile.Readline
		   TrackStr =  Mid(strNextLine,10, len(strNextLine) - 9)
		   strNextLine = objTextFile.Readline
		   GenreStr =  Mid(strNextLine,10, len(strNextLine) - 9)
		   strNextLine = objTextFile.Readline
		   YearStr =  Mid(strNextLine,10, len(strNextLine) - 9)
		   strNextLine = objTextFile.Readline
		   FileStr =  Mid(strNextLine,10, len(strNextLine) - 9)
		   strNextLine = objTextFile.Readline
		   CountStr =  Mid(strNextLine,10, len(strNextLine) - 9)
		   sleep 5
		   SetpropertyValue "Jukebox.Jukebox - Import Messages", "Adding song: " & TitleStr
		   SetpropertyValue "Jukebox.Jukebox - Library Song Count", CountStr
		   
		   'this checks for a bug in id3.exe 
		   If Left(GenreStr,1)  = "(" Then
			  str_pos = InStr(GenreStr, ")")
			  If str_pos <> 0 Then
				 data = Split(GenreStr,")")
				 GenreStr = data(1)
			  End if
		   End If
		   'SqlStr = "Select file from songs where file = '" & FileStr & "'"
			SqlStr = su.Sprintf("Select id, file from songs where file = %Nq",FileStr)
			Set r = objDB.Execute(SqlStr)
		   ' Only insert if new song is found  -- Otherwise, Relationships will be compromised
		   If r.count < 1 Then
			  objDB.Execute su.Sprintf("INSERT INTO songs (title,artist,album,track,genre,year,file) VALUES (%Nq,%Nq,%Nq,%Nq,%Nq,%Nq,%Nq)",TitleStr,ArtistStr,AlbumStr,TrackStr,GenreStr,YearStr,FileStr)
		   Else
			  For Row = 1 To r.count
				song_id = r(row)("id")
				objDB.Execute su.Sprintf("Update songs set title = %Nq ,artist = %Nq, album = %Nq, track = %Nq, genre = %Nq, year = %Nq, file = %Nq where id = %Nq",TitleStr,ArtistStr,AlbumStr,TrackStr,GenreStr,YearStr,FileStr,song_id)
			  Next
		   End If
		   Set r = Nothing
		End if
	Loop


	SetpropertyValue "Jukebox.Jukebox - Import Messages", "Cleaning up Orphaned Songs"
	'Clean Up  - Delete all entries where a file is not found
	Set objFile = CreateObject("Scripting.FileSystemObject")

	Set r = objDB.Execute("select * from songs")
	For Row = 1 To r.Count
	   If objFile.FileExists(r(row)("file")) = False Then
		  objDB.Execute("delete from songs where id = " & r(row)("id"))
	   End If
	Next 
	Set r = Nothing


	SetpropertyValue "Jukebox.Jukebox - Import Messages", "Cleaning up Orphaned Saved Playlist"
	'Clean Up  - Delete Orphaned saved playlists 
	Set r = objDB.Execute("select songid from savedplaylists")
	For Row = 1 To r.Count
	   SetpropertyValue "Jukebox Script.Error", r(Row)("songid")    
	   Set s = objDB.Execute("select id from songs where id = " & r(Row)("songid"))
	   If s.Count = 0 Then
		  objDB.Execute("delete from savedplaylists where songid = " & r(Row)("songid"))
		  
	   End If
	Next

	SetpropertyValue "Jukebox.Jukebox - Import Messages", "Cleaning up Orphaned Songs in Song Queues"
	'Clean Up  - Delete Orphaned songs in songqueue 
	Set r = objDB.Execute("select songid from songqueue")
	For Row = 1 To r.Count
	   SetpropertyValue "Jukebox Script.Error", r(Row)("songid")    
	   Set s = objDB.Execute("select id from songs where id = " & r(Row)("songid"))
	   If s.Count = 0 Then
		  objDB.Execute("delete from songqueue where songid = " & r(Row)("songid"))
	   End If
	Next
	 
	Set r = Nothing
	Set s = Nothing

	objDB.Execute su.Sprintf("insert into genre (genre) select distinct genre from songs")

	Set oShell = Nothing
	objTextFile.Close
	Set objFile = Nothing
	Set objTextFile = Nothing
	Set objFSO = Nothing

	SetpropertyValue "Jukebox.Jukebox - Import Messages", "Populating Genre List"
	PopulateGenreList
	sleep 500
	SetpropertyValue "Jukebox.Jukebox - Import Messages", "Populating Name List"
	Call PopulateNameList("False", 0, "Remote-1")
	Call PopulateNameList("False", 1, "Remote-2")
	Call PopulateNameList("False", 2, "Remote-3")

	sleep 500
	PopulatePlayList  0, 0
	PopulatePlayList  0, 1
	PopulatePlayList  0, 2
	SetpropertyValue "Jukebox.Jukebox - Import Messages", "Populating Play List"
	SetpropertyValue "Jukebox.Jukebox - Import Flag", "Stopped"
	SetpropertyValue "Jukebox.Jukebox - Import Messages", ""
	
End Sub


Function GetRemoteNumber(RemoteName)
	Dim NoMoreRemotes, RemoteCount, RemoteNumber
	RemoteCount = 0
	NoMoreRemotes = 0
	RemoteNumber = ""
	Do Until NoMoreRemotes = 1
		If GetPropertyValue ("Remote-" + CStr(RemoteCount+1) + ".Remote Name") <> "* error *" Then
		   RemoteCount=RemoteCount + 1
			If GetPropertyValue ("Remote-" + CStr(RemoteCount) + ".Remote Name") = RemoteName Then
				RemoteNumber = "Remote-" + CStr(RemoteCount)	
			End if
		Else
			NoMoreRemotes = 1
		End if
	Loop
	GetRemoteNumber = RemoteNumber
End Function


' ********************** Global Functions: Needs to Be in all Subscriber Files****************************
Sub SendSubscriberMessage(SubscriberID, SubscriberMessage)
	do 
		If GetPropertyValue("Subscriber-" & CStr(SubscriberID) & ".DispatchMessage") = "Idle" Then
			'SetPropertyValue "System.Debug", "Got Subscriber"
			'Perform A Crude Form of Record Locking 
			BlockSubscriber(SubscriberID)	
			SetpropertyValue "Subscriber-" & Cstr(SubscriberID) & ".DispatchMessage", SubscriberMessage
			UnBlockSubscriber(SubscriberID) 
			Exit Do
		Else
			Sleep GetRandomNumber (1,100)
		End if
	Loop

End Sub

Sub BlockSubscriber(SubscriberID) 
	SetModeState "Subscriber-" & CStr(SubscriberID), "Inactive"
End Sub	

Sub UnBlockSubscriber(SubscriberID) 
	SetModeState "Subscriber-" & CStr(SubscriberID), "Active"
End Sub	

Function GetRandomNumber (min,max)
	Randomize
	GetRandomNumber = Int((max-min+1)*Rnd+min)
End Function

' EOF ********************** Needs to Be in all Subscriber Files****************************



