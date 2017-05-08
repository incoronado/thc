Option Explicit
'On Error Resume Next
' Message Format MessageHandlerID.Class.Source.Priority.Command:Param1:Param2

Dim Action
Dim CurrentTimeData
Dim OldTimeData
Dim SleepVar
Dim ReceivedData
Dim OldReceivedData
Dim CurrentAVPowerSetting
Dim OldAVPowerSetting
Dim CurrentMatrixData, CurrentAMPM, CurrentTime
Dim OldMatrixData
Dim month_day
Dim abbrev_month
Dim OldRDSPS1, OldRDSRT1, OldRDSPS2, OldRDSRT2, OldSelectedSource, OldAlbum
Dim RoomMenuStatus, FloorMenuStatus, FloorMenuSelected
Dim HousebotLocation, DOSHousebotLocation 

DOSHousebotLocation = "C:\Progra~2\Housebot\"
HousebotLocation = "C:\Program Files (x86)\housebot\"
RoomMenuStatus = 0
FloorMenuStatus = 0

' All = 0, FirstFloor = 1, SecondFloor = 2
FloorMenuSelected = 1

' Set Menu Auto Close Timeout
SetpropertyValue "Menu Timer.Sleep Time", "Days=00, Hours=00, Minutes=00, Seconds=04, MilliSeconds=00"
HBRemoteList
'Set su = CreateObject("newObjects.utilctls.StringUtilities")

'-------------------------------------------------------
'- Main: Checks Received Data and Handles Send Data ----
'-------------------------------------------------------

SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
Do
	Sleep SleepVar
  	Action = GetPropertyValue ("System.Action")
	If Action <> "Idle" Then
		SystemCommand(Action)
		SetpropertyValue "System.Action", "Idle"
	End If
	'Sleep SleepVar	
Loop


Sub SystemCommand(Action)
	'menu.floor.room
	dim a,b,x, bcount, ar_floors, ar_first_rooms, ar_second_rooms, SourceVar, CurrentVolume, FrequencyVar, i, objDB, su, r, StationList, Row, sUrl, sRequest, PlayerType, SelectedJukebox
    'SetpropertyValue "UPB Script.Debug", UPBLightID
    'MsgBox "Ready Command"
	a=split(Action,".")
	sUrl = "http://192.168.1.4:8080/jsonrpc"
	Select Case a(0)
		' System Class Message Received
		Case "System"
		b=split(a(3),":")
		bcount = Ubound(b) + 1
		'System.GalaxyTabA1.10.SelectZone:1
			Select Case b(0)
				Case "DailyTasks"
					'System.GalaxyTabA1.10.DailyTasks
					SetPropertyValue "System.Abbreviated DOW", ucase(Mid(GetPropertyValue("System Time.Day Of Week"),1,3))
					abbrev_month = split(GetpropertyValue("System Time.Date"),"/")
					If abbrev_month(0) = 1 Then
						SetPropertyValue "System.Abbreviated Month", "JAN"
					Elseif abbrev_month(0) = 2 Then	
						SetPropertyValue "System.Abbreviated Month", "FEB"
					Elseif abbrev_month(0) = 3 Then
						SetPropertyValue "System.Abbreviated Month", "MAR"		
					Elseif abbrev_month(0) = 4 Then	
						SetPropertyValue "System.Abbreviated Month", "APR"
					Elseif abbrev_month(0) = 5 Then	
						SetPropertyValue "System.Abbreviated Month", "MAY"
					Elseif abbrev_month(0) = 6 Then	
						SetPropertyValue "System.Abbreviated Month", "JUN"
					Elseif abbrev_month(0) = 7 Then	
						SetPropertyValue "System.Abbreviated Month", "JUL"
					Elseif abbrev_month(0) = 8 Then	
						SetPropertyValue "System.Abbreviated Month", "AUG"
					Elseif abbrev_month(0) = 9 Then	
						SetPropertyValue "System.Abbreviated Month", "SEP"
					Elseif abbrev_month(0) = 10 Then	
						SetPropertyValue "System.Abbreviated Month", "OCT"
					Elseif abbrev_month(0) = 11 Then	
						SetPropertyValue "System.Abbreviated Month", "NOV"
					Elseif abbrev_month(0) = 12 Then	
						SetPropertyValue "System.Abbreviated Month", "DEC"
					End if
					month_day = split(GetpropertyValue("System Time.Date"),"/")
					SetPropertyValue "System.Month Day", month_day(1)
				Case "ZoneSourceOn"
					'System.GalaxyTabA1.10.ZoneSourceOn:Zone:Source
					'System.GalaxyTabA1.10.ZoneSourceOn:5:1
					'Source: 1=Jukebox 1;2=Tuner 1;3=Cable TV;4=Tuner 2; 5=Jukebox 2; 6=Apple TV
					'Zone: 1=Living Room; 5=Master Bedroom; 6=Patio
					ZoneSourceOn b(1), b(2)
				Case "ZoneOff"
					ZoneOff b(1)
				Case "UnblankRemote"
					'System.GalaxyTabA1.10.UnblankRemote
					SetWhichRemotesToControl a(1)
					UnBlankRemoteScreen
				Case "BlankRemote"
					'System.GalaxyTabA1.10.BlankRemote
					SetWhichRemotesToControl a(1)
					BlankRemoteScreen	
				
				Case "SelectZone"
				'System.GalaxyTabA1.10.SelectZone:1
					SelectZone GetRemoteNumber(a(1)), b(1)
				Case "ToggleZonePower"
					'System.GalaxyTabA1.10.ToggleZonePower
					ToggleZonePower(GetRemoteNumber(a(1)))
				Case "VolumeUp"
					'System.GalaxyTabA1.10.VolumeUp
					VolumeUp(GetRemoteNumber(a(1)))	
				Case "VolumeDown"
					'System.GalaxyTabA1.10.VolumeDown
					VolumeDown(GetRemoteNumber(a(1)))
				Case "ChannelDown"
					'System.GalaxyTabA1.10.ChannelDown
					ChannelDown(GetRemoteNumber(a(1)))
				Case "ChannelUp"
					'System.GalaxyTabA1.10.ChannelUp
					ChannelUp(GetRemoteNumber(a(1)))
				Case "SetZoneSource"
					'System.GalaxyTabA1.10.SetZoneSource:2
					SetZoneSource GetRemoteNumber(a(1)), b(1)
				Case "ToggleLinkSource"
					'System.GalaxyTabA1.10.ToggleLinkSource
					ToggleLinkSource(GetRemoteNumber(a(1)))
				Case "OpenSelectedJukeboxPanel"
					'System.GalaxyTabA1.10.OpenSelectedJukeboxPanel
					SelectedJukebox=GetPropertyValue(GetRemoteNumber(a(1)) & ".Selected Jukebox")
					SelectPanel a(1), "Jukebox", "Jukebox " & SelectedJukebox, 0
				Case "SelectRemotePanel"	
					'System.GalaxyTabA1.10.SelectRemotePanel
					SelectRemotePanel(GetRemoteNumber(a(1)))
				Case "SelectNextRemote"
					'System.GalaxyTabA1.10.SelectNextRemote
					SelectNextRemote(GetRemoteNumber(a(1)))
				Case "SelectMusicPanel"
					'System.GalaxyTabA1.10.SelectMusicPanel
					SelectMusicPanel(GetRemoteNumber(a(1)))
				Case "ToggleNavigatorPanel"
					'System.Desktop.10.ToggleNavigatorPanel
					ToggleNavigatorPanel(GetRemoteNumber(a(1)))
				Case "SelectPanel" 
					'System.GalaxyTabA1.10.SelectPanel:Jukebox:Jukebox1:5
					SelectPanel a(1), b(1), b(2), b(3)	
				Case "ClosePanel"
					'System.GalaxyTabA1.10.ClosePanel:Navigator
					ClosePanel GetRemoteNumber(a(1)), b(1)
				Case "SelectNavigatorPanel"
					'System.GalaxyTabA1.10.SelectMusicPanel
					SelectSelectNavigatorPanelPanel(GetRemoteNumber(a(1)))	
				Case "SelectSprinklersPanel"
					'System.GalaxyTabA1.10.SelectMusicPanel
					OpenTHCRemotePanel "Sprinklers", GetRemoteNumber(a(1))
				Case "TurnAllAVOff"
					'System.GalaxyTabA1.10.TurnAllAVOff:A
					TurnAllAVOff b(1)
				Case "AVOn"
					'System.GalaxyTabA1.10.AVOn:CableTVOn:B
					AVOn b(1), b(2)
				Case "AVOn2"
					'System.GalaxyTabA1.10.AVOn2:Cable TV:Master Bedroom
					AVOn2 b(1), b(2)
				Case "AVOff2"
					'System.GalaxyTabA1.10.AVOff2:Master Bedroom
					AVOff2 b(1)	
				Case "ToggleBedroomTV"	
					SetPropertyValue "USBUIRT.Westinghouse Remote", "Power"
				Case "TogglePlayer"
					TogglePlayer b(1)
					'System.GalaxyTabA1.10.TogglePlayer:1
				Case "PlayerStatus" 
					PlayerStatus (GetRemoteNumber(a(1)))
				Case "CallerID"
					'System.GalaxyTabA1.10.CallerID
					HTTPPost sURL, "{""jsonrpc"":""2.0"",""method"":""GUI.ShowNotification"",""params"":{""title"":""Incoming Call"",""message"":""" & GetPropertyValue("CallerID.Last Caller") & ":" & GetPropertyValue("CallerID.Last Phone Number") & """,""displaytime"":20000},""id"":1}"
					SetWhichRemotesToControl a(1) 
					UnBlankRemoteScreen
					
					SelectPanel a(1), "Caller ID", "", 10	
					
					
				Case "PanasonicTVOn"

				
				Case "PanasonicTVOff"

				Case "WestinghouseTVPower" 
					'System.GalaxyTabA1.10.WestinghouseTVPower
					SetpropertyValue "USBUIRT.Westinghouse Remote", "Power"
					
				Case "XBMC"
					Select Case lcase(b(1))
						Case "guide"
							'System.GalaxyTabA1.10.XBMC:Guide
							HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""GUI.ActivateWindow"",""params"":{""window"":""tvguide""} ,""id"" :1}"
						Case "input"
							'POST Request to send.
							' System.GalaxyTabA1.10.XBMC:Input:Home
							sRequest = "{""jsonrpc"": ""2.0"", ""method"": """ & b(1) & "." & b(2) & """, ""id"" :1}"
							HTTPPost sURL, sRequest																																																																   
						Case "player"
							Select Case lcase(b(2))
								Case "play"
									'System.GalaxyTabA1.10.XBMC:Player:Play
									PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
									HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.PlayPause"",""params"":{""playerid"":" & PlayerType & ",""play"":true} ,""id"" :1}"
								Case "pause"
									'System.GalaxyTabA1.10.XBMC:Player:Pause
									PlayerType = HTTPPost(sURL,"{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
									HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.PlayPause"",""params"":{""playerid"":" & PlayerType & ",""play"":false} ,""id"" :1}"
								Case "ff"
									'System.GalaxyTabA1.10.XBMC:Player:FF
									PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
									HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.SetSpeed"",""params"":{""playerid"":" & PlayerType & ",""speed"":4} ,""id"" :1}"
								Case "rw"
									'System.GalaxyTabA1.10.XBMC:Player:RW
									PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
									HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.SetSpeed"",""params"":{""playerid"":" & PlayerType & ",""speed"":-4} ,""id"" :1}"
								Case "stop"
									'System.GalaxyTabA1.10.XBMC:Player:Stop
									PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
									HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.Stop"",""params"":{""playerid"":" & PlayerType & "} ,""id"" :1}"
								Case "next"
									'System.GalaxyTabA1.10.XBMC:Player:Next
									PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
									HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GoTo"",""params"":{""playerid"":" & PlayerType & ", ""to"": ""next""} ,""id"" :1}"
								Case "previous"
									'System.GalaxyTabA1.10.XBMC:Player:Previous
									PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
									HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GoTo"",""params"":{""playerid"":" & PlayerType & ", ""to"": ""previous""} ,""id"" :1}"
							End Select
						Case "callerid"
							'System.GalaxyTabA1.10.XBMC:CallerID
							HTTPPost sURL, "{""jsonrpc"":""2.0"",""method"":""GUI.ShowNotification"",""params"":{""title"":""Incoming Call"",""message"":""" & GetPropertyValue("CallerID.Last Caller") & ":" & GetPropertyValue("CallerID.Last Phone Number") & """,""displaytime"":20000},""id"":1}"		
					End Select	
				
				Case "AllMediaOff"
					'System.GalaxyTabA1.10.AllMediaOff
					SetpropertyValue "System.Matrix Zone " & a(1) & " Power State", "Off"
			
					sleep 50
					If a(1) = "A" Then
						SetpropertyValue "USBUIRT.Panasonic Remote", "Power Off"
						SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOff"
					End if	
					If ((GetpropertyValue("System.Matrix Zone A Power State") = "Off") And (GetpropertyValue("System.Matrix Zone B Power State") = "Off") And (GetpropertyValue("System.Matrix Zone C Power State") = "Off") And (GetpropertyValue("System.Matrix Zone D Power State") = "Off")) Then
						SetpropertyValue "HDMI Matrix Script.Action", "PowerOff"
						SetpropertyValue "USBUIRT.Sony BD", "Power Off"
						SetpropertyValue "USBUIRT.Direct TV Remote", "Power Off"
					End if		
			End Select
			
		Case "SetRes"
			SetRes
		Case "UpdateTimeDigits"
			UpdateTimeDigits
		Case "TogglePlayer"
			TogglePlayer "GalaxyProTab1"
		Case "PlayerStatus" 
			PlayerStatus "GalaxyProTab1"
		Case "Select Next Remote"
			SelectNextRemote "GalaxyProTab1"
		Case "Toggle Link Source"	
			ToggleLinkSource "GalaxyProTab1"
		Case "CallerID"
			HTTPPost sURL, "{""jsonrpc"":""2.0"",""method"":""GUI.ShowNotification"",""params"":{""title"":""Incoming Call"",""message"":""" & GetPropertyValue("CallerID.Last Caller") & ":" & GetPropertyValue("CallerID.Last Phone Number") & """,""displaytime"":20000},""id"":1}"
	    Case "Update Tuner 1 Station Image"
			  SetPropertyValue "Tuner 1 Settings.Now Playing Image", "config\themes\THC\Radio Stations\" & GetPropertyValue("Tuner 1 Settings.Tuner FREQ") & ".png"	
              If CInt(GetPropertyValue("GalaxyProTab1" & " System.Selected Source")) = 2 Then
				SetPropertyValue "GalaxyProTab1" & " System.Radio Station Frequency", DispFREQ(GetPropertyValue("Tuner 1 Settings.Tuner FREQ"))
				SetPropertyValue "GalaxyProTab1" & " System.Notification Text 1", ""
				SetPropertyValue "GalaxyProTab1" & " System.Notification Text 2", ""
		  	  End if  
		Case "Update Tuner 2 Station Image"
			SetPropertyValue "Tuner 2 Settings.Now Playing Image", "config\themes\THC\Radio Stations\" & GetPropertyValue("Tuner 2 Settings.Tuner FREQ") & ".png"	
			If CInt(GetPropertyValue("GalaxyProTab1" & " System.Selected Source")) = 4 Then	
				SetPropertyValue "GalaxyProTab1" & " System.Radio Station Frequency", DispFREQ(GetPropertyValue("Tuner 2 Settings.Tuner FREQ"))
				SetPropertyValue "GalaxyProTab1" & " System.Notification Text 1", ""
				SetPropertyValue "GalaxyProTab1" & " System.Notification Text 2", ""
			End if
	  Case "Audio"
			select case a(2)
				Case "SelectZone"
					SetPropertyValue "GalaxyProTab1" & " System.Selected Source", GetpropertyValue("Multiroom Audio Settings.Zone " + a(3) + " Source")
					SetPropertyValue "GalaxyProTab1" & " System.Selected Zone", CStr(a(3))
					SetPropertyValue "GalaxyProTab1" & " System.Selected Zone Power", GetpropertyValue("Multiroom Audio Settings.Zone " + a(3) + " Power")
					SetPropertyValue "GalaxyProTab1" & " System.Selected Zone Volume", GetpropertyValue("Multiroom Audio Settings.Zone " + a(3) + " Volume")
				Case "SetZoneSource"
					SetPropertyValue "GalaxyProTab1" & " System.Selected Source", a(3)
				    If GetPropertyValue ("GalaxyProTab1" &" System.Link Source") = "Linked" Then
						SetpropertyValue "Multiroom Audio Script.Action", "Source." &  GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & "." & a(3)
					End If			
					If CInt(GetPropertyValue("GalaxyProTab1" & " System.Selected Source")) = 2 Then
						SetpropertyValue "GalaxyProTab1" & " System.Selected Tuner", "1"
						SetPropertyValue "GalaxyProTab1" & " System.Radio Station Frequency", DispFREQ(GetPropertyValue("Tuner 1 Settings.Tuner FREQ"))
						SetPropertyValue "GalaxyProTab1" & " System.Notification Text 1", ""
						SetPropertyValue "GalaxyProTab1" & " System.Notification Text 2", ""
					ElseIf CInt(GetPropertyValue("GalaxyProTab1" & " System.Selected Source")) = 4 Then	
						SetpropertyValue "GalaxyProTab1" & " System.Selected Tuner", "2"
						SetPropertyValue "GalaxyProTab1" & " System.Radio Station Frequency", DispFREQ(GetPropertyValue("Tuner 2 Settings.Tuner FREQ"))
						SetPropertyValue "GalaxyProTab1" & " System.Notification Text 1", ""
						SetPropertyValue "GalaxyProTab1" & " System.Notification Text 2", ""
					Else
					    SetpropertyValue "GalaxyProTab1" & " System.Selected Tuner", ""
					End If	
				Case "ToggleZonePower"
					If GetPropertyValue("Multiroom Audio Settings.Zone " & GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & " Power") = "On" Then
						SetPropertyValue "MRA Ignore Receive.Running", "Yes"
						SetPropertyValue "GalaxyProTab1" & " System.Selected Zone Power", "Off"
						SetPropertyValue "Multiroom Audio Settings.Zone " & GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & " Power", "Off"	
						SetpropertyValue "Multiroom Audio Script.Action", "Off." &  GetPropertyValue("GalaxyProTab1" & " System.Selected Zone")
					Else
						SourceVar = GetPropertyValue("GalaxyProTab1" & " System.Selected Source")
						SetPropertyValue "MRA Ignore Receive.Running", "Yes"
						SetPropertyValue "GalaxyProTab1" & " System.Selected Zone Power", "On"
						SetPropertyValue "Multiroom Audio Settings.Zone " & GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & " Power", "On"
						SetpropertyValue "Multiroom Audio Script.Action", "On." &  GetPropertyValue("GalaxyProTab1" & " System.Selected Zone")
						sleep 50
						'SetpropertyValue "Multiroom Audio Settings.Debug", "Source." &  GetPropertyValue(a(1) & " System.Selected Zone") & "." & SourceVar
						SetpropertyValue "Multiroom Audio Script.Action", "Source." &  GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & "." & SourceVar
					End If
				Case "VolumeUp"	
				'SetPropertyValue "Nav Panel.Running", "Yes"
				SetPropertyValue "MRA Ignore Receive.Running", "Yes"
					CurrentVolume=CInt(GetPropertyValue("Multiroom Audio Settings.Zone " & GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & " Volume"))
					If CurrentVolume < 38 Then 
						SetPropertyValue "GalaxyProTab1" & " System.Selected Zone Volume", CurrentVolume+1
						SetPropertyValue "Multiroom Audio Settings.Zone " & GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & " Volume", CurrentVolume+1
						SetpropertyValue "Multiroom Audio Script.Action", "Vol." &  GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & "." & CurrentVolume+1
					End If	
				Case "VolumeDown"
					'SetPropertyValue "Nav Panel.Running", "Yes"
					SetPropertyValue "MRA Ignore Receive.Running", "Yes"
					CurrentVolume=CInt(GetPropertyValue("Multiroom Audio Settings.Zone " & GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & " Volume"))
					If CurrentVolume > 0 Then 
						SetPropertyValue "GalaxyProTab1" & " System.Selected Zone Volume", CurrentVolume-1
						SetPropertyValue "Multiroom Audio Settings.Zone " & GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & " Volume", CurrentVolume-1
						SetpropertyValue "Multiroom Audio Script.Action", "Vol." &  GetPropertyValue("GalaxyProTab1" & " System.Selected Zone") & "." & CurrentVolume-1
					End If
				Case "ChannelUp"
					If CInt(GetPropertyValue("GalaxyProTab1" & " System.Selected Source")) = 2 Then 
						SetpropertyValue "Tuner 1.Tuner Settings", "Tune Up" 
					ElseIf CInt(GetPropertyValue(a(1) & " System.Selected Source")) = 4 Then
						SetpropertyValue "Tuner 2.Tuner Settings", "Tune Up" 	
					End If
				Case "ChannelDown"
				    If CInt(GetPropertyValue(a(1) & " System.Selected Source")) = 2 Then 
						SetpropertyValue "Tuner 1.Tuner Settings", "Tune Down" 
					ElseIf CInt(GetPropertyValue(a(1) & " System.Selected Source")) = 4 Then
						SetpropertyValue "Tuner 2.Tuner Settings", "Tune Down" 	
					End If
				Case "PopulateStations"
					' Open Database
					Set objDB  = CreateObject("newObjects.sqlite3.dbutf8")
					objDB.Open(HousebotLocation & "config\scripts\songs.mlf")
					objDB.BusyTimeout=1000
					objDB.AutoType = True
					objDB.TypeInfoLevel = 4
					
					Set su = CreateObject("newObjects.utilctls.StringUtilities")
					StationList = ""
					Set r = objDB.Execute("select * from stations")
					For Row = 1 To r.Count
						 StationList = StationList & r(Row)("id") & vbTAB & r(Row)("frequency") & vbTAB & r(Row)("call-letters") & vbTAB & r(Row)("format") & vbLF
						
					Next 	
					
					SetPropertyValue "System.Tuner 1 Radio Stations", Left(StationList,Len(StationList)-1)
					sleep 5
					SetPropertyValue "System.Tuner 2 Radio Stations", Left(StationList,Len(StationList)-1)
					objDB.Close
					Set objDB = Nothing
					Set su = Nothing
	
				Case "Change Station"
					Set objDB  = CreateObject("newObjects.sqlite3.dbutf8")
					objDB.Open(HousebotLocation & "config\scripts\songs.mlf")
					objDB.BusyTimeout=1000
					objDB.AutoType = True
					objDB.TypeInfoLevel = 4
					Dim FreqAR
					StationList = ""
					Set r = objDB.Execute("select * from stations")
					
					For Row = 1 To r.Count
					    if  CInt(GetPropertyValue("System.Tuner " & a(1) & " Selected Radio Station")) = r(Row)("preset") Then
						   'SetpropertyValue "System Script.Debug", r(Row)("preset") & ", " &  GetPropertyValue("System.Tuner " & a(1) & " Selected Radio Station")
							StationList = StationList & "*S-" & r(Row)("id") & vbTAB & r(Row)("frequency") & vbTAB & r(Row)("call-letters") & vbTAB & r(Row)("format") & vbLF
							SetPropertyValue "System.Radio Station Call Letters", r(Row)("call-letters")
							SetPropertyValue "System.Radio Station Frequency", r(Row)("frequency") & "MHz"
							FreqAR = split(r(Row)("frequency"),".")
							SetPropertyValue "Tuner " & a(1) & ".Frequency Direct", CStr(FreqAR(0)) & "_" & Right(CStr(FreqAR(1)) & "0", 2) 
										
						Else
							StationList = StationList & r(Row)("id") & vbTAB & r(Row)("frequency") & vbTAB & r(Row)("call-letters") & vbTAB & r(Row)("format") & vbLF
						End if	
					Next 	
					SetPropertyValue "System.Tuner " & a(1) & " Radio Stations", Left(StationList,Len(StationList)-1)
					objDB.Close
					Set objDB = Nothing
					Set su = Nothing
			End Select
		Case "XBMC"
				select Case a(1) 
					Case "Guide"
						HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""GUI.ActivateWindow"",""params"":{""window"":""tvguide""} ,""id"" :1}"
					Case "Input"
						'POST Request to send.
						sRequest = "{""jsonrpc"": ""2.0"", ""method"": """ & a(1) & "." & a(2) & """, ""id"" :1}"
						HTTPPost sURL, sRequest																																																																   
					Case "Player"
						Select Case a(2) 
							Case "Play"
								PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
								'setpropertyValue "System.Debug", PlayerType
								'SetpropertyValue "System.Debug", PlayerType
								HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.PlayPause"",""params"":{""playerid"":" & PlayerType & ",""play"":true} ,""id"" :1}"
							Case "Pause"
								PlayerType = HTTPPost(sURL,"{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
								HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.PlayPause"",""params"":{""playerid"":" & PlayerType & ",""play"":false} ,""id"" :1}"
							Case "FF"
								PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
								HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.SetSpeed"",""params"":{""playerid"":" & PlayerType & ",""speed"":4} ,""id"" :1}"
							Case "RW"
								PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
								HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.SetSpeed"",""params"":{""playerid"":" & PlayerType & ",""speed"":-4} ,""id"" :1}"
							Case "Stop"
							    PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
								HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.Stop"",""params"":{""playerid"":" & PlayerType & "} ,""id"" :1}"
							Case "Next"
							    PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
								HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GoTo"",""params"":{""playerid"":" & PlayerType & ", ""to"": ""next""} ,""id"" :1}"
							Case "Previous"
							    PlayerType = HTTPPost(sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GetActivePlayers"",""id"" :1}")
								HTTPPost sURL, "{""jsonrpc"": ""2.0"", ""method"": ""Player.GoTo"",""params"":{""playerid"":" & PlayerType & ", ""to"": ""previous""} ,""id"" :1}"
						End Select
				End Select
			
		' Configuration Command
		Case "Abbreviated DOW"
			SetPropertyValue "System.Abbreviated DOW", ucase(Mid(GetPropertyValue("System Time.Day Of Week"),1,3))
			
		Case "Abbreviated Month"	
			abbrev_month = split(GetpropertyValue("System Time.Date"),"/")
			If abbrev_month(0) = 1 Then
				SetPropertyValue "System.Abbreviated Month", "JAN"
			Elseif abbrev_month(0) = 2 Then	
				SetPropertyValue "System.Abbreviated Month", "FEB"
			Elseif abbrev_month(0) = 3 Then
				SetPropertyValue "System.Abbreviated Month", "MAR"		
			Elseif abbrev_month(0) = 4 Then	
				SetPropertyValue "System.Abbreviated Month", "APR"
			Elseif abbrev_month(0) = 5 Then	
				SetPropertyValue "System.Abbreviated Month", "MAY"
			Elseif abbrev_month(0) = 6 Then	
				SetPropertyValue "System.Abbreviated Month", "JUN"
			Elseif abbrev_month(0) = 7 Then	
				SetPropertyValue "System.Abbreviated Month", "JUL"
			Elseif abbrev_month(0) = 8 Then	
				SetPropertyValue "System.Abbreviated Month", "AUG"
			Elseif abbrev_month(0) = 9 Then	
				SetPropertyValue "System.Abbreviated Month", "SEP"
			Elseif abbrev_month(0) = 10 Then	
				SetPropertyValue "System.Abbreviated Month", "OCT"
			Elseif abbrev_month(0) = 11 Then	
				SetPropertyValue "System.Abbreviated Month", "NOV"
			Elseif abbrev_month(0) = 12 Then	
				SetPropertyValue "System.Abbreviated Month", "DEC"
			End if
		Case "OpenRoomMenu"

		
		Case "Month Day"
			month_day = split(GetpropertyValue("System Time.Date"),"/")
			SetPropertyValue "System.Month Day", month_day(1)
		
		Case "DirectTVOn"
			SetpropertyValue "System.Matrix Zone " & a(1) & " Power State", "On"
			If a(1) = "A" Then
			    Sleep 150
				SetpropertyValue "USBUIRT.Panasonic Remote", "Power On"
			Elseif a(1) = "B" Then
			    'SetpropertyValue "Multiroom Audio Script.Action", "On.5"
				'Sleep 50
				SetpropertyValue "USBUIRT.Westinghouse Remote", "Power"
				Sleep 50
				SendSubscriberMessage 1,"MRA.IRRemote.10.On:5"
				'SetpropertyValue "Subscriber-1.DispatchMessage", "MRA.IRRemote.10.On:5"
			End if
			Sleep 50	
			SetpropertyValue "USBUIRT.Direct TV Remote", "Power On"                               
			If ((GetpropertyValue("Yamaha V2600 Settings.AV Power Master") = "Off") And (a(1) = "A")) Then
				SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOn"
			End if
			
			If GetpropertyValue("HDMI Matrix Settings.HDMI Power State") = "Off" Then
				SetpropertyValue "HDMI Matrix Script.Action", "PowerOn"
					Sleep 50
			End if
			If GetpropertyValue("System.Zone " & a(1) & " Selected Source") <> "Direct TV" Then
				SetpropertyValue "HDMI Matrix Script.Action", "Set" & a(1) & "4"
			End if
		
		Case "XBMCOn"
			SetpropertyValue "System.Matrix Zone " & a(1) & " Power State", "On"
		    If a(1) = "A" Then
				SetpropertyValue "USBUIRT.Panasonic Remote", "Power On"
				
			End if	
			If ((GetpropertyValue("Yamaha V2600 Settings.AV Power Master") = "Off")  And (a(1) = "A")) Then
				SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOn"
			End if
			If GetpropertyValue("HDMI Matrix Settings.HDMI Power State") = "Off" Then
				SetpropertyValue "HDMI Matrix Script.Action", "PowerOn"
				Sleep 50
			End if
			If GetpropertyValue("System.Zone " & a(1) & " Selected Source") <> "HA Server" Then
				SetpropertyValue "HDMI Matrix Script.Action", "Set" & a(1) & "1"
			End if
	
			
			
		Case "AppleTVOn"
			SetpropertyValue "System.Matrix Zone " & a(1) & " Power State", "On"
			If a(1) = "A" Then
				SetpropertyValue "USBUIRT.Panasonic Remote", "Power On"
			End if	
			
			If ((GetpropertyValue("Yamaha V2600 Settings.AV Power Master") = "Off") And (a(1) = "A")) Then
				SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOn"
			End if
			If GetpropertyValue("HDMI Matrix Settings.HDMI Power State") = "Off" Then
				SetpropertyValue "HDMI Matrix Script.Action", "PowerOn"
				Sleep 50
			End if
			If GetpropertyValue("System.Zone " & a(1) & " Selected Source") <> "Apple TV" Then
				SetpropertyValue "HDMI Matrix Script.Action", "Set" & a(1) & "2"
			End if
		
		Case "BlueRayOn"
			SetpropertyValue "System.Matrix Zone " & a(1) & " Power State", "On"
			If a(1) = "A" Then
				SetpropertyValue "USBUIRT.Panasonic Remote", "Power On"
			End if
			SetpropertyValue "USBUIRT.Sony BD", "Power On"
			If ((GetpropertyValue("Yamaha V2600 Settings.AV Power Master") = "Off") And (a(1) = "A")) Then
				SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOn"
			End if
			If GetpropertyValue("HDMI Matrix Settings.HDMI Power State") = "Off" Then
				SetpropertyValue "HDMI Matrix Script.Action", "PowerOn"
				Sleep 50
			End if
			If GetpropertyValue("System.Zone " & a(1) & " Selected Source") <> "DVD" Then
				SetpropertyValue "HDMI Matrix Script.Action", "Set" & a(1) & "3"
			End if
			
		Case "AllMediaOff"
			SetpropertyValue "System.Matrix Zone " & a(1) & " Power State", "Off"
			
			sleep 50
			If a(1) = "A" Then
				SetpropertyValue "USBUIRT.Panasonic Remote", "Power Off"
				SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOff"
			End if	
			If ((GetpropertyValue("System.Matrix Zone A Power State") = "Off") And (GetpropertyValue("System.Matrix Zone B Power State") = "Off") And (GetpropertyValue("System.Matrix Zone C Power State") = "Off") And (GetpropertyValue("System.Matrix Zone D Power State") = "Off")) Then
				SetpropertyValue "HDMI Matrix Script.Action", "PowerOff"
				SetpropertyValue "USBUIRT.Sony BD", "Power Off"
				SetpropertyValue "USBUIRT.Direct TV Remote", "Power Off"
			End if	
			
		Case "Close Remote Panel"
			CloseRemotePanel("Remote")
		Case "Select Home Panel"
			CloseAllPanels "GalaxyProTab1"
			OpenRemotePanel("Home")
		Case "Select Remote Panel"
		    CloseAllPanels "GalaxyProTab1"
			SelectCurrentRemote "GalaxyProTab1"
		Case "Select Movies Panel"
			CloseAllPanels "GalaxyProTab1"
			OpenRemotePanel("Movies")
		Case "Open Music Library Panel"	
			CloseAllPanels "GalaxyProTab1"
			OpenRemotePanel("Music Library")
		Case "Close Tuner Panel"
			CloseAllPanels "GalaxyProTab1"
			CloseRemotePanel("Tuner")	
		Case "Select Music Panel"
		     
		Case "Select Tuner Panel"
		    CloseAllPanels "GalaxyProTab1"
			SetWhichRemotesToControl("")
			If CInt(GetPropertyValue("GalaxyProTab1" & " System.Selected Source")) = 2 Then
				OpenRemotePanelAndSetContext "Tuner", "Tuner 1"
			ElseIf CInt(GetPropertyValue("GalaxyProTab1" & " System.Selected Source")) = 4 Then
				OpenRemotePanelAndSetContext "Tuner", "Tuner 2"
			Else
				OpenRemotePanelAndSetContext "Tuner", "Tuner " & GetPropertyValue("GalaxyProTab1" & " System.Selected Tuner")
			End if	
			SetWhichRemotesToControl("")

		Case "Open Playlist Panel"
		    SetWhichRemotesToControl("")
			OpenRemotePanel("Playlist")
			SetWhichRemotesToControl("")
		Case "Select Lights Panel"
		    CloseAllPanels "GalaxyProTab1"
			SetWhichRemotesToControl("")
			OpenRemotePanel("Lights")
			SetWhichRemotesToControl("")
		Case "Close Playlist Panel"	
		    
			SetWhichRemotesToControl("")
			CloseRemotePanel("Playlist")
			CloseRemotePanel("Keyboard")
			CloseAllPanels "GalaxyProTab1"
			SetWhichRemotesToControl("")
		Case "Close All Panels"
			CloseAllPanels "GalaxyProTab1"
	End Select
		 
End Sub

Sub TurnAllAVOff(Zone)
	SetpropertyValue "System.Matrix Zone " & Zone & " Power State", "Off"
	
	sleep 50
	If Zone = "A" Then
		SetpropertyValue "USBUIRT.Panasonic Remote", "Power Off"
		SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOff"
	End if	
	If ((GetpropertyValue("System.Matrix Zone A Power State") = "Off") And (GetpropertyValue("System.Matrix Zone B Power State") = "Off") And (GetpropertyValue("System.Matrix Zone C Power State") = "Off") And (GetpropertyValue("System.Matrix Zone D Power State") = "Off")) Then
		SetpropertyValue "HDMI Matrix Script.Action", "PowerOff"
		SetpropertyValue "USBUIRT.Sony BD", "Power Off"
		SetpropertyValue "USBUIRT.Direct TV Remote", "Power Off"
	End if	
End Sub


Sub SelectMusicPanel(Remote)
	CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
	SetWhichRemotesToControl(GetPropertyValue(Remote & ".Remote Name"))  
	If CInt(GetPropertyValue(Remote & ".Selected Source")) = 1 Then
		OpenRemotePanelAndSetContext "Jukebox", "Jukebox 1"
	ElseIf CInt(GetPropertyValue(Remote & ".Selected Source")) = 5 Then
		OpenRemotePanelAndSetContext "Jukebox", "Jukebox 2"
	Else
		OpenRemotePanelAndSetContext "Jukebox", "Jukebox " & GetPropertyValue(Remote & ".Selected Jukebox")
	End if
	SetWhichRemotesToControl("")
End Sub


Sub ToggleNavigatorPanel(Remote)

	If GetPropertyValue(Remote & ".Nav Menu Status") = "Inactive" Then
    	SetPropertyValue Remote & ".Nav Menu Status", "Active"
    	SelectPanel GetPropertyValue(Remote & ".Remote Name"), "Navigator","", 0
    Else
    	SetPropertyValue Remote & ".Nav Menu Status", "Inactive"
    	ClosePanel Remote, "Navigator"
    End If
End Sub


Sub SelectPanel(Remote, Panel, Context, TimerInSeconds)
	 Dim hours, minutes, seconds, IntSeconds
	 
	If Panel <> "Navigator" Then
	   CloseAllPanels Remote
	End if	
	
	SetWhichRemotesToControl Remote
	If Context = "" Then
	    SetPropertyValue "System.Debug", Panel
		OpenRemotePanel Panel
	Else
		OpenRemotePanelAndSetContext Panel, Context
	End if

	
	If TimerinSeconds > 0 Then
		' calculates whole hours (like a div operator)
		hours = TimerInSeconds \ 3600
		' calculates the remaining number of seconds
		intSeconds = TimerInSeconds Mod 3600
		' calculates the whole number of minutes in the remaining number of seconds
		minutes = intSeconds \ 60
		' calculates the remaining number of seconds after taking the number of minutes
		seconds = intSeconds Mod 60
		hours = Right("00" & CStr(hours),2)
		minutes = Right("00" & CStr(minutes),2)
		seconds = Right("00" & CStr(seconds),2)
	    SetpropertyValue "System.Command Timer 1", "System." & Remote & ".10.ClosePanel:" & Panel 
		SetpropertyValue "Command Timer 1.Sleep Time", "Days=00, Hours=" & hours & ", Minutes=" & minutes & ", Seconds=" & seconds & ", MilliSeconds=00"
		SetpropertyValue "Command Timer 1.Running", "Yes"
	End If
	SetWhichRemotesToControl("")
End Sub

Sub ClosePanel(Remote, Panel)
	'CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
	SetWhichRemotesToControl(GetPropertyValue(Remote & ".Remote Name"))
	CloseRemotePanel Panel
	SetWhichRemotesToControl("")
End Sub

Sub ZoneSourceOn (Zone, Source)
	SendSubscriberMessage 1, "MRA.System.10.On:" &  Zone
	'SetpropertyValue "Subscriber-1.DispatchMessage", "MRA.System.10.On:" &  Zone
	'Sleep 100
	SendSubscriberMessage 1, "MRA.System.10.Source:" &  Zone & ":" & Source
	'SetpropertyValue "Subscriber-1.DispatchMessage", "MRA.System.10.Source:" &  Zone & ":" & Source
End Sub

Sub ZoneOff (Zone)
	SendSubscriberMessage 1, "MRA.System.10.Off:" &  Zone
	'SetpropertyValue "Subscriber-1.DispatchMessage", "MRA.System.10.Off:" &  Zone
End Sub




Sub TogglePlayer (PlayerNo)
	If GetPropertyValue("Music Player " & PlayerNo & ".Status") = "Playing" Then
	   SetPropertyValue("Music Player " & PlayerNo & ".Status"), "Paused" 
	Else
	   SetPropertyValue("Music Player " & PlayerNo & ".Status"), "Playing" 
	End if
End Sub 

Sub PlayerStatus (RemoteName)
	If CInt(GetPropertyValue(RemoteName & " System.Selected Source")) = 1 Then
	    SetPropertyValue (RemoteName) & " System.Player Status", GetPropertyValue("Music Player 1.Status")
	ElseIf CInt(GetPropertyValue(RemoteName & " System.Selected Source")) = 5 Then
		SetPropertyValue RemoteName & " System.Player Status", GetPropertyValue("Music Player 2.Status")
	End if
 
End Sub

Function DispFREQ(Frequency)
	DispFREQ = mid(Frequency,1,2) & " " & mid(Frequency,3,len(Frequency)-4) & "."  & right(Frequency,2)
End Function

Sub UpdateTime
	SetpropertyValue "System.EpochTime", DateDiff("s", "01/01/1970 00:00:00", now())
End Sub

Sub GreatRoomPowerOff
	If GetPropertyValue("Yamaha V2600 Settings.AV Power Master") = "On" Then
		SetpropertyValue "HDMI Matrix Script.Action", "PowerOn"
	Else
		SetpropertyValue "HDMI Matrix Script.Action", "PowerOff"
		SetpropertyValue "USBUIRT.Direct TV Remote", "Power Off"
	End if
End Sub

Sub UpdateTimeDigits
		CurrentAMPM = UCase(Right(GetPropertyValue("System Time.TimeAndDate"),2))

		If GetpropertyValue ("System.AMPM") <> CurrentAMPM Then
			SetpropertyValue "System.AMPM", CurrentAMPM
			SetpropertyValue "System.AM PM Digit 1", mid(CurrentAMPM,1,1)
			SetpropertyValue "System.AM PM Digit 2", mid(CurrentAMPM,2,1)
		End If
		Currenttime = Right("0" & GetPropertyValue("System Time.Time Without Seconds"),8)
		If mid(CurrentTime,1,1) = 0 Then
			SetpropertyValue "System.Time Digit 1", mid(CurrentTime,1,1)
		Else
		    SetpropertyValue "System.Time Digit 1", mid(CurrentTime,1,1)   
		End if
		SetpropertyValue "System.Time Digit 2", mid(CurrentTime,2,1)
		SetpropertyValue "System.Time Digit 3", mid(CurrentTime,4,1)
		SetpropertyValue "System.Time Digit 4", mid(CurrentTime,5,1)
		
End Sub


Sub TimedTask(TimeInSeconds,Action) 
	StartTimer = GetpropertyValue("System.EpochTime")
End Sub


Function HTTPPost(sUrl, sRequest)
	dim objXmlHttpMain
	Set objXmlHttpMain = CreateObject("Microsoft.XMLHTTP") 
	'on error resume next 
	objXmlHttpMain.open "POST", sUrl, False
	'objXmlHttpMain.setRequestHeader "Authorization", "kodi kodi"
	objXmlHttpMain.setRequestHeader "Content-Type", "application/json" 
	objXmlHttpMain.send(sRequest)
	HTTPPost = PlayerID(objXmlHttpMain.responseText)
	set objXmlHttpMain = nothing
	'set objJSONDoc = nothing 
End Function

Sub CloseAllPanels (RemoteName)
    SetWhichRemotesToControl(RemoteName) 
	'CloseRemotePanel("Navigator")
    CloseRemotePanel("Remote")
    CloseRemotePanel("Music Library")
	CloseRemotePanel("Nav")
	CloseRemotePanel("Home")
	CloseRemotePanel("Tuner")
	CloseRemotePanel("Jukebox")
	CloseRemotePanel("Keyboard")
	CloseRemotePanel("Genre")
	CloseRemotePanel("Playlist")
	CloseRemotePanel("Lights")
	CloseRemotePanel("Movies")
	CloseRemotePanel("Sprinklers")
	sleep 50
	SetWhichRemotesToControl("") 

End Sub

Sub SelectCurrentRemote (RemoteName)
CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
	CloseAllPanels GetPropertyValue(RemoteName & ".Remote Name")
	SetWhichRemotesToControl(GetPropertyValue(RemoteName & ".Remote Name")) 
	OpenRemotePanelAndSetContext "Remote", GetPropertyValue(RemoteName & ".Selected Remote")	
	SetWhichRemotesToControl("") 
End Sub 

Sub SelectNextRemote (RemoteName)
	'CloseAllPanels GetPropertyValue(RemoteName & ".Remote Name")
	SetWhichRemotesToControl(GetPropertyValue(RemoteName & ".Remote Name")) 
	If GetPropertyValue(RemoteName & ".Selected Remote") = "Blu-Ray" Then
		SetPropertyValue RemoteName & ".Selected Remote", "Apple TV"
		OpenRemotePanelAndSetContext "Remote", "Apple TV"
	ElseIf GetPropertyValue(RemoteName & ".Selected Remote") = "Apple TV" Then
		SetPropertyValue RemoteName & ".Selected Remote", "Kodi"
		OpenRemotePanelAndSetContext "Remote", "Kodi"
	Elseif GetPropertyValue(RemoteName & ".Selected Remote") = "Kodi" Then
		SetPropertyValue RemoteName & ".Selected Remote", "Cable TV"
		OpenRemotePanelAndSetContext "Remote", "Cable TV"
	Else
		SetPropertyValue RemoteName & ".Selected Remote", "Blu-Ray"
		OpenRemotePanelAndSetContext "Remote", "Blu-Ray"
	End If	
	SetWhichRemotesToControl("") 
End Sub 


Function PlayerID(json)
	'SetpropertyValue "System.Debug", json
	'SetpropertyValue "System.Debug", instr(json,"video")
	  
	If json = "{""id"":1,""jsonrpc"":""2.0"",""result"":[{""playerid"":1,""type"":""video""}]}" Then
	   PlayerID = "1"
	Elseif json = "{""id"":1,""jsonrpc"":""2.0"",""result"":[{""playerid"":0,""type"":""audio""}]}" Then
	   PlayerID = "0"
	End if	

End Function

Sub SetRes
	Dim CmdStr, oShell
	Set oShell = CreateObject( "WScript.Shell" )
	CmdStr = "NirCmd setdisplay 1024 768 32"
	oShell.Run CmdStr, 0, True
	sleep 100
	CmdStr = "NirCmd setdisplay 1280 1024 32"
	oShell.Run CmdStr, 0, True
	
	
End Sub

Sub ToggleLinkSource(RemoteName)
    If GetPropertyValue (RemoteName & ".Link Source") = "Linked" Then
	   SetPropertyValue RemoteName & ".Link Source", "Unlinked"
	Else
	   SetPropertyValue RemoteName & ".Link Source", "Linked"
	End If
End Sub

Sub VolumeDown(Remote)
	Dim CurrentVolume
	SetPropertyValue "MRA Ignore Receive.Running", "Yes"

	CurrentVolume=CInt(GetPropertyValue("Multiroom Audio Settings.Zone " & GetPropertyValue(Remote & ".Selected Zone") & " Volume"))
		
	If CurrentVolume > 0 Then 
		'SetPropertyValue Remote & ".Selected Zone Volume", CurrentVolume-1
		SetPropertyValue "Multiroom Audio Settings.Zone " & GetPropertyValue(Remote & ".Selected Zone") & " Volume", CurrentVolume-1
		'SetpropertyValue "Multiroom Audio Script.Action", "Vol." &  GetPropertyValue(Remote & ".Selected Zone") & "." & CurrentVolume-1
		SendSubscriberMessage 1, "MRA." & Remote & ".10.Vol:" &  GetPropertyValue(Remote & ".Selected Zone") & ":" & CurrentVolume-1
		'SetpropertyValue "Subscriber-1.DispatchMessage", "MRA." & Remote & ".10.Vol:" &  GetPropertyValue(Remote & ".Selected Zone") & ":" & CurrentVolume-1
	End If
End Sub

Function GetZoneNumber(ZoneName)

	Dim NoMoreZones, ZoneCount, ZoneNumber
	ZoneCount = 0
	NoMoreZones = 0
	ZoneNumber = ""
	Do Until NoMoreZones = 1
		If GetPropertyValue ("Multiroom Audio Settings.Zone " + CStr(RemoteCount+1) + " TV") <> "* error *" Then
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

Function GetSourceNumber(SourceName)

End Function



Sub VolumeUp(Remote)
	Dim CurrentVolume
	SetPropertyValue "MRA Ignore Receive.Running", "Yes"
	CurrentVolume=CInt(GetPropertyValue("Multiroom Audio Settings.Zone " & GetPropertyValue(Remote & ".Selected Zone") & " Volume"))
	If CurrentVolume < 38 Then 
		'SetPropertyValue Remote & ".Selected Zone Volume", CurrentVolume+1
		SetPropertyValue "Multiroom Audio Settings.Zone " & GetPropertyValue(Remote & ".Selected Zone") & " Volume", CurrentVolume+1
		SendSubscriberMessage 1, "MRA." & Remote & ".10.Vol:" &  GetPropertyValue(Remote & ".Selected Zone") & ":" & CurrentVolume+1
		'SetpropertyValue "Subscriber-1.DispatchMessage", "MRA." & Remote & ".10.Vol:" &  GetPropertyValue(Remote & ".Selected Zone") & ":" & CurrentVolume+1
	End If	
End Sub

Sub SetZoneSource(Remote,Source)
	SetPropertyValue Remote & ".Selected Source", CStr(Source)
	SetPropertyValue "Multiroom Audio Settings.Zone " & GetPropertyValue(Remote & ".Selected Zone") & " Source", CStr(Source)
    If GetPropertyValue (Remote & ".Link Source") = "Linked" Then
		
		If GetPropertyValue(Remote & ".Selected Zone") = 5 Then
			If CInt(Source) = 3 Then
				AVOn "CableTVOn", "B"
			ElseIf CInt(Source) = 6 Then
				AVOn "AppleTVOn", "B"
			Else
				SendSubscriberMessage 1,"MRA.System.10.zonesource:5:" & CStr(Source)
			End if
		ElseIf GetPropertyValue(Remote & ".Selected Zone") = 1 Then
		
		Else
			SendSubscriberMessage 1, "MRA.System.10.source:" & GetPropertyValue(Remote & ".Selected Zone") & ":" & CStr(Source)
		End if
	End If	
	
	If CInt(Source) = 1 Then
		SetpropertyValue Remote & ".Selected Jukebox", "1"
		CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
		SetWhichRemotesToControl(GetPropertyValue(Remote & ".Remote Name"))
		OpenRemotePanelAndSetContext "Jukebox", "Jukebox 1"
		SetWhichRemotesToControl("")
		
		
	ElseIf CInt(Source) = 5 Then
		SetpropertyValue Remote & ".Selected Jukebox", "2"
		CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
		SetWhichRemotesToControl(GetPropertyValue(Remote & ".Remote Name"))
		OpenRemotePanelAndSetContext "Jukebox", "Jukebox 2"
		SetWhichRemotesToControl("")
	
	ElseIf CInt(Source) = 2 Then
		SetpropertyValue Remote & ".Selected Tuner", "1"
		SetPropertyValue Remote & ".Radio Station Frequency", DispFREQ(GetPropertyValue("Tuner 1 Settings.Tuner FREQ"))
		SetPropertyValue Remote & ".Notification Text 1", ""
		SetPropertyValue Remote & ".Notification Text 2", ""
		CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
		SetWhichRemotesToControl(GetPropertyValue(Remote & ".Remote Name"))
		OpenRemotePanelAndSetContext "Tuner", "Tuner 1"
		SetWhichRemotesToControl("")
	ElseIf CInt(Source) = 4 Then	
		SetpropertyValue Remote & ".Selected Tuner", "2"
		SetPropertyValue Remote & ".Radio Station Frequency", DispFREQ(GetPropertyValue("Tuner 2 Settings.Tuner FREQ"))
		SetPropertyValue Remote & ".Notification Text 1", ""
		SetPropertyValue Remote & ".Notification Text 2", ""
		CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
		SetWhichRemotesToControl(GetPropertyValue(Remote & ".Remote Name"))
		OpenRemotePanelAndSetContext "Tuner", "Tuner 2"
		SetWhichRemotesToControl("")
	ElseIf CInt(Source) = 3 Then
		SetpropertyValue Remote & ".Selected Remote", "Cable TV"
		CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
		SetWhichRemotesToControl(GetPropertyValue(Remote & ".Remote Name"))
		OpenRemotePanelAndSetContext "Remote", "Cable TV"
		SetWhichRemotesToControl("")
	ElseIf CInt(Source) = 6 Then
		SetpropertyValue Remote & ".Selected Remote", "Apple TV"
		CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
		SetWhichRemotesToControl(GetPropertyValue(Remote & ".Remote Name"))
		OpenRemotePanelAndSetContext "Remote", "Apple TV"
		SetWhichRemotesToControl("")	
	Else
	    SetpropertyValue Remote & ".Selected Tuner", ""
	End If	
	
End Sub


Sub ToggleZonePower(Remote)
	Dim SourceVar			    
	If GetPropertyValue("Multiroom Audio Settings.Zone " & GetPropertyValue(Remote & ".Selected Zone") & " Power") = "On" Then
		'SetPropertyValue "MRA Ignore Receive.Running", "Yes"
		SetPropertyValue Remote & ".Selected Zone Power", "Off"
		SetPropertyValue "Multiroom Audio Settings.Zone " & GetPropertyValue(Remote & ".Selected Zone") & " Power", "Off"
        If GetPropertyValue(Remote & ".Selected Zone") = 5 Then
			SetpropertyValue "USBUIRT.Westinghouse Remote", "Power"
		End If	
		SendSubscriberMessage 1,"MRA." & Remote & ".10.Off:" &  GetPropertyValue(Remote & ".Selected Zone")
		'SetpropertyValue "Subscriber-1.DispatchMessage", "MRA." & Remote & ".10.Off:" &  GetPropertyValue(Remote & ".Selected Zone")
	Else
		SourceVar = GetPropertyValue(Remote & ".Selected Source")
		'SetPropertyValue "MRA Ignore Receive.Running", "Yes"
		SetPropertyValue Remote & ".Selected Zone Power", "On"
		SetPropertyValue "Multiroom Audio Settings.Zone " & GetPropertyValue(Remote & ".Selected Zone") & " Power", "On"
		'SetpropertyValue "Multiroom Audio Script.Action", "On." &  GetPropertyValue(Remote & ".Selected Zone")
		SendSubscriberMessage 1, "MRA." & Remote & ".10.On:" &  GetPropertyValue(Remote & ".Selected Zone")
		'SetpropertyValue "Subscriber-1.DispatchMessage", "MRA." & Remote & ".10.On:" &  GetPropertyValue(Remote & ".Selected Zone")
		If GetPropertyValue(Remote & ".Selected Zone") = 5 Then
			SetpropertyValue "USBUIRT.Westinghouse Remote", "Power"
		End If
		sleep 50
		SendSubscriberMessage 1, "MRA." & Remote & ".10.Source:" &  GetPropertyValue(Remote & ".Selected Zone") & ":" & SourceVar
		'SetpropertyValue "Subscriber-1.DispatchMessage", "MRA." & Remote & ".10.Source:" &  GetPropertyValue(Remote & ".Selected Zone") & ":" & SourceVar
		'SetpropertyValue "Subscriber-1.DispatchMessage", "MRA." & Remote & ".10.Off:" &  GetPropertyValue(Remote & ".Selected Zone")
	End If
End Sub


Sub SelectZone (Remote,ZoneNo)
	SetPropertyValue Remote & ".Selected Source", GetpropertyValue("Multiroom Audio Settings.Zone " + CStr(ZoneNo) + " Source")
	SetPropertyValue Remote & ".Selected Zone", CStr(ZoneNo)
	SetPropertyValue Remote & ".Selected Zone Power", GetpropertyValue("Multiroom Audio Settings.Zone " + CStr(ZoneNo) + " Power")
	SetPropertyValue Remote & ".Selected Zone Volume", GetpropertyValue("Multiroom Audio Settings.Zone " + CStr(ZoneNo) + " Volume")
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

Sub ChannelUp(Remote)
	If CInt(GetPropertyValue(Remote & ".Selected Source")) = 2 Then 
		SetpropertyValue "Tuner 1.Tuner Settings", "Tune Up" 
	ElseIf CInt(GetPropertyValue(Remote & ".Selected Source")) = 4 Then
		SetpropertyValue "Tuner 2.Tuner Settings", "Tune Up"
	ElseIf CInt(GetPropertyValue(Remote & ".Selected Source")) = 3 Then
		SetpropertyValue "USBUIRT.Cable TV Remote", "Channel +"
	End If
End Sub
			
Sub	ChannelDown(Remote)
    If CInt(GetPropertyValue(Remote & ".Selected Source")) = 2 Then 
		SetpropertyValue "Tuner 1.Tuner Settings", "Tune Down" 
	ElseIf CInt(GetPropertyValue(Remote & ".Selected Source")) = 4 Then
		SetpropertyValue "Tuner 2.Tuner Settings", "Tune Down" 
	ElseIf CInt(GetPropertyValue(Remote & ".Selected Source")) = 3 Then
		SetpropertyValue "USBUIRT.Cable TV Remote", "Channel -"
		
	End If
End Sub

Sub SelectRemotePanel(Remote)
	CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
	SetWhichRemotesToControl(GetPropertyValue(Remote & ".Remote Name"))
	OpenRemotePanelAndSetContext "Remote", GetPropertyValue(Remote & ".Selected Remote")	
	SetWhichRemotesToControl("")
End Sub


Sub OpenTHCRemotePanel(Panel, Remote)
	CloseAllPanels GetPropertyValue(Remote & ".Remote Name")
	SetWhichRemotesToControl(GetPropertyValue(Remote & ".Remote Name"))
	OpenRemotePanelAndSetContext Panel, GetPropertyValue(Remote & ".Selected Remote")	
	SetWhichRemotesToControl("")
End Sub

' Test
Sub AVOn2 (Source, Zone)
	'Rewrite of AVOn that takes advantage of HB propery settings.  Core of the Audio Video Logic
    'Check to see if TV is applicable
    
	' Turn Video On If Applicable to Zone and Source   
	If ((CInt(GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV")) = 1) And (CInt(GetPropertyValue("Multiroom Audio Settings.Source " & CStr(SourceName2ID(Source)) & " TV")) = 1))  Then
		SetpropertyValue "System.Matrix Zone " & VideoZoneName2Alpha(Zone) & " Power State", "On"
		'Turn On Video Matrix If Off
		If GetpropertyValue("HDMI Matrix Settings.HDMI Power State") = "Off" Then
			SetpropertyValue "HDMI Matrix Script.Action", "PowerOn"
			Sleep 50
		End if
		' Set the Vidoe Matrix to the Source and Zone
		SetpropertyValue "HDMI Matrix Script.Action", "Set" & VideoZoneName2Alpha(Zone) & VideoSourceName2Number(Source) 
		'SetpropertyValue "System.Debug", "Set" & VideoZoneName2Alpha(Zone) & VideoSourceName2Number(Source) 
		
		'Send TV On Command
		if Trim(GetPropertyValue("Multiroom Audio Settings.Source " & CStr(SourceName2ID(Source)) & " TV")) = "1" Then
			If GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV Status") = "0" Then		
				SendSubscriberMessage 1, GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV On Command")
				SetPropertyValue "Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV Status", "1"
			End if	
		Else	
			If GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV Status") = "1" Then		
				SendSubscriberMessage 1, GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV Off Command")	
				SetPropertyValue "Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV Status", "0"
			End If	
		End if
	End if

	'Turn Blue Ray Player On 
	If Source = "Blu-Ray" Then
		SendSubscriberMessage 1, "IR.System.10.SonyBluRay:On"
	End if

	If Zone = "Living Room" Then

		'Send Yamaha Command
		'If (GetpropertyValue("Yamaha V2600 Settings.AV Power Master") = "Off" Then
			SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOn"	
		'End if

		Select Case Source
			Case "Cable TV"
				SendSubscriberMessage 1, "Receiver.RS232.10.Input:Cbl/Sat"
			Case "Jukebox 1"
				SendSubscriberMessage 1, "Receiver.RS232.10.Input:dvr/vcr2"
				SendSubscriberMessage 1, "Russound.System.10.ZonePower:" & Zone & ":on"
				SendSubscriberMessage 1, "Russound.System.10.ZoneSource:" & Zone & ":" & Source
			Case "Jukebox 2"
				SendSubscriberMessage 1, "Receiver.RS232.10.Input:dvr/vcr2"
				SendSubscriberMessage 1, "Russound.System.10.ZonePower:" & Zone & ":on"
				SendSubscriberMessage 1, "Russound.System.10.ZoneSource:" & Zone & ":" & Source
			Case "Streamer"
				SendSubscriberMessage 1, "Receiver.RS232.10.Input:dvr/vcr2"
				SendSubscriberMessage 1, "Russound.System.10.ZonePower:" & Zone & ":on"
				SendSubscriberMessage 1, "Russound.System.10.ZoneSource:" & Zone & ":" & Source
			Case "Radio"			
				SendSubscriberMessage 1, "Receiver.RS232.10.Input:dvr/vcr2"
				SendSubscriberMessage 1, "Russound.System.10.ZonePower:" & Zone & ":on"
				SendSubscriberMessage 1, "Russound.System.10.ZoneSource:" & Zone & ":" & Source
			Case "Blu-Ray"
				SendSubscriberMessage 1, "Receiver.RS232.10.Input:Cbl/Sat"
			Case "Kodi"
				SendSubscriberMessage 1, "Receiver.RS232.10.Input:Cbl/Sat"
			Case "Apple TV"
				SendSubscriberMessage 1, "Receiver.RS232.10.Input:Cbl/Sat"
		End Select
	Else
		SendSubscriberMessage 1, "Russound.System.10.ZonePower:" & Zone & ":on"
		SendSubscriberMessage 1, "Russound.System.10.ZoneSource:" & Zone & ":" & Source
	End if	




End Sub

Sub AVOff2 (Zone)
	Dim BluRayOn, a, item
	BluRayOn = 0

	SetpropertyValue "System.Matrix Zone " & VideoZoneName2Alpha(Zone) & " Power State", "Off"

	If ((GetpropertyValue("Yamaha V2600 Settings.AV Power Master") = "On") And (Zone = "Living Room")) Then
		SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOff"
	Else
		SendSubscriberMessage 1, "Russound.System.10.ZonePower:" & Zone & ":off"
	End if	

    ' Turn Video Off If Applicable to Zone
    If CInt(GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV")) = 1  Then
		SetpropertyValue "System.Matrix Zone " & VideoZoneName2Alpha(Zone) & " Power State", "Off"
	End if	

	If GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV Status") <> "0" Then
		if Trim(GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV Off Command")) <> "" Then
			SendSubscriberMessage 1, GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV Off Command")
		End If	
		SetPropertyValue "Multiroom Audio Settings.Zone " & CStr(ZoneName2ID(Zone)) & " TV Status", 0
	End if


	a=Array("A","B","C","D")
 	For each item in a
    	If GetPropertyValue("System.Matrix Zone " & item & " Power State") = "On" Then
    		If GetPropertyValue("System.Zone " & item & " Selected Source") = "Blu-Ray" Then
    			BluRayOn = 1
    		End If
    	End if
 	Next

 	If BluRayOn = 1 Then
 		SendSubscriberMessage 1, "IR.System.10.SonyBluRay:Off"
 	End If

End Sub


Function VideoZoneName2Alpha(ZoneName) 
	Dim ZoneAlpha, ReturnValue, item
	ZoneAlpha=Array("A","B","C","D")
 	ReturnValue=""
 	For each item in ZoneAlpha
    	If lcase(ZoneName) = lcase(GetPropertyValue("System.Matrix Zone " & item & " Name")) Then
    		ReturnValue = item
    	End if
 	Next
 	VideoZoneName2Alpha = ReturnValue
End Function

Function VideoSourceName2Number(SourceName) 
	Dim ZoneAlpha, ReturnValue, item
	ZoneAlpha=Array("1","2","3","4")
 	ReturnValue=""
 	For each item in ZoneAlpha						
    	If lcase(SourceName) = lcase(GetPropertyValue("System.Matrix Source " & item & " Name")) Then
    		ReturnValue = item
    	End if
 	Next
 	VideoSourceName2Number = ReturnValue
End Function


Function SourceName2ID(SourceName) 
	Dim SourceCount, NoMoreSources, SourceNo
	SourceNo=0
	NoMoreSources = 0
	SourceCount = 0
	Do Until NoMoreSources = 1
		If GetPropertyValue("Multiroom Audio Settings.Source " & CStr(SourceCount+1) & " Name") <> "* error *" Then
		   If GetPropertyValue("Multiroom Audio Settings.Source " & CStr(SourceCount+1) & " Name") = SourceName Then
		   		SourceNo = SourceCount + 1
		   End if
		   SourceCount=SourceCount + 1  
		Else
			NoMoreSources = 1
		End if
	Loop
	'SetPropertyValue "Multiroom Audio Settings.Debug 2", SourceNo
	SourceName2ID = SourceNo
End Function

Function ZoneName2ID(ZoneName) 
	Dim ZoneCount, NoMoreZones, ZoneNo
	ZoneNo=0
	NoMoreZones = 0
	ZoneCount = 0
	Do Until NoMoreZones = 1
		If GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneCount+1) &  " Name") <> "* error *" Then
		   If GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneCount+1) & " Name") = ZoneName Then
		   		ZoneNo = ZoneCount + 1
		   End if
		   ZoneCount=ZoneCount + 1  
		Else
			NoMoreZones = 1
		End if
	Loop
	SetPropertyValue "Multiroom Audio Settings.Debug 2", ZoneNo
	ZoneName2ID = ZoneNo
End Function



Sub AVOn (AVFunction, VideoZone)
	Select Case lcase(AVFunction)
		Case "cabletvon"
			SetpropertyValue "System.Matrix Zone " & VideoZone & " Power State", "On"
			If VideoZone = "A" Then
				SetpropertyValue "USBUIRT.Panasonic Remote", "Power On"
			Elseif VideoZone = "B" Then
				SendSubscriberMessage 1, "Russound.Desktop.10.ZonePower:Master Bedroom:on"
				sleep 50
				SendSubscriberMessage 1, "Russound.Desktop.10.ZoneSource:Master Bedroom:Cable TV"
				'Allow 5MS to Queue Message
				Sleep 50
			End if
			Sleep 50	
			SetpropertyValue "USBUIRT.Direct TV Remote", "Power On"                               
			If ((GetpropertyValue("Yamaha V2600 Settings.AV Power Master") = "Off") And (VideoZone = "A")) Then
				SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOn"
			End if
			
			If GetpropertyValue("HDMI Matrix Settings.HDMI Power State") = "Off" Then
				SetpropertyValue "HDMI Matrix Script.Action", "PowerOn"
					Sleep 50
			End if
			If GetpropertyValue("System.Zone " & VideoZone & " Selected Source") <> "Direct TV" Then
				SetpropertyValue "HDMI Matrix Script.Action", "Set" & VideoZone & "4"
			End if
			
			SetpropertyValue "System.Zone " & VideoZone & " Selected Source", "Direct TV"
		
		Case "xbmcon"
			SetpropertyValue "System.Matrix Zone " & VideoZone & " Power State", "On"
		    If VideoZone = "A" Then
				SetpropertyValue "USBUIRT.Panasonic Remote", "Power On"
				Elseif VideoZone = "B" Then
				SendSubscriberMessage 1, "Russound.GalaxyTabA1.10.ZonePower:Master Bedroom:On"
				sleep 50
				SendSubscriberMessage 1, "Russound.Desktop.10.ZoneSource:Master Bedroom:Kodi"	
			End if	
			
			If ((GetpropertyValue("Yamaha V2600 Settings.AV Power Master") = "Off")  And (VideoZone = "A")) Then
				SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOn"
			End if
			If GetpropertyValue("HDMI Matrix Settings.HDMI Power State") = "Off" Then
				SetpropertyValue "HDMI Matrix Script.Action", "PowerOn"
				Sleep 50
			End if
			If GetpropertyValue("System.Zone " & VideoZone & " Selected Source") <> "HA Server" Then
				SetpropertyValue "HDMI Matrix Script.Action", "Set" & VideoZone & "1"
			End if
			
		Case "appletvon"
			SetpropertyValue "System.Matrix Zone " & VideoZone & " Power State", "On"
			If VideoZone = "A" Then
				SetpropertyValue "USBUIRT.Panasonic Remote", "Power On"
			Elseif VideoZone = "B" Then
				SendSubscriberMessage 1, "Russound.GalaxyTabA1.10.ZonePower:Master Bedroom:On"
				sleep 50
				SendSubscriberMessage 1, "Russound.Desktop.10.ZoneSource:Master Bedroom:Apple TV"
			End if	
			
			If ((GetpropertyValue("Yamaha V2600 Settings.AV Power Master") = "Off") And (VideoZone = "A")) Then
				SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOn"
			End if
			If GetpropertyValue("HDMI Matrix Settings.HDMI Power State") = "Off" Then
				SetpropertyValue "HDMI Matrix Script.Action", "PowerOn"
				Sleep 50
			End if
			If GetpropertyValue("System.Zone " & VideoZone & " Selected Source") <> "Apple TV" Then
				SetpropertyValue "HDMI Matrix Script.Action", "Set" & VideoZone & "2"
			End if
			SetpropertyValue "System.Zone " & VideoZone & " Selected Source", "Apple TV"
		
		Case "bluerayon"
			SetpropertyValue "System.Matrix Zone " & VideoZone & " Power State", "On"
			If VideoZone = "A" Then
				SetpropertyValue "USBUIRT.Panasonic Remote", "Power On"
			Elseif VideoZone = "B" Then
				SendSubscriberMessage 1, "Russound.GalaxyTabA1.10.ZonePower:Master Bedroom:On"
				sleep 50
				SendSubscriberMessage 1, "Russound.Desktop.10.ZoneSource:Master Bedroom:Blu-Ray"

			End if
			SetpropertyValue "USBUIRT.Sony BD", "Power On"
			If ((GetpropertyValue("Yamaha V2600 Settings.AV Power Master") = "Off") And (VideoZone = "A")) Then
				SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOn"
			End if
			If GetpropertyValue("HDMI Matrix Settings.HDMI Power State") = "Off" Then
				SetpropertyValue "HDMI Matrix Script.Action", "PowerOn"
				Sleep 50
			End if
			If GetpropertyValue("System.Zone " & VideoZone & " Selected Source") <> "DVD" Then
				SetpropertyValue "HDMI Matrix Script.Action", "Set" & VideoZone & "3"
			End if
			
		Case "allmediaoff"
			SetpropertyValue "System.Matrix Zone " & VideoZone & " Power State", "Off"
			
			sleep 50
			If VideoZone = "A" Then
				SetpropertyValue "USBUIRT.Panasonic Remote", "Power Off"
				SetpropertyValue "Yamaha V2600 Settings.Action", "MasterPowerOff"
			End if	
			If ((GetpropertyValue("System.Matrix Zone A Power State") = "Off") And (GetpropertyValue("System.Matrix Zone B Power State") = "Off") And (GetpropertyValue("System.Matrix Zone C Power State") = "Off") And (GetpropertyValue("System.Matrix Zone D Power State") = "Off")) Then
				SetpropertyValue "HDMI Matrix Script.Action", "PowerOff"
				SetpropertyValue "USBUIRT.Sony BD", "Power Off"
				SetpropertyValue "USBUIRT.Direct TV Remote", "Power Off"
			End if	
		End Select
End Sub

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

Sub HBRemoteList
	Dim NoMoreRemotes, RemoteCount, HBRList
	HBRList = ""
	RemoteCount = 0
	NoMoreRemotes = 0
	Do Until NoMoreRemotes = 1
		If ((GetPropertyValue ("Remote-" + CStr(RemoteCount+1) + ".Remote Name") <> "* error *") And (GetPropertyValue ("Remote-" + CStr(RemoteCount+1) + ".Remote Type") = "HB")) Then
			HBRList = HBRList & CStr(RemoteCount+1) & "," 	
			RemoteCount=RemoteCount + 1
		Else
			NoMoreRemotes = 1
		End if
		
	Loop
	SetPropertyValue "System.HBRemoteCount", CInt(RemoteCount)
	SetPropertyValue "System.HBRemoteList", Left(HBRList,Len(HBRList)-1) 
End Sub

' EOF ********************** Needs to Be in all Subscriber Files****************************


Sub ParsonMovieJSON (jsonSTR)
	Dim fso, json, str, o, i
	Set json = New VbsJson
	Set o = json.Decode(jsonSTR)
	WScript.Echo o("Result")("moviedetails")("art")("fanart")
	WScript.Echo o("Result")("moviedetails")("art")("poster")
	For Each i In o("Result")("moviedetails")("director")
		WScript.Echo i
	Next
	For Each i In o("Result")("moviedetails")("genre")
		WScript.Echo i
	Next
	WScript.Echo o("Result")("moviedetails")("label")
	WScript.Echo o("Result")("moviedetails")("movieid")
	WScript.Echo o("Result")("moviedetails")("mpaa")
	WScript.Echo o("Result")("moviedetails")("label")
	WScript.Echo o("Result")("moviedetails")("plot")
	WScript.Echo o("Result")("moviedetails")("playcount")
	WScript.Echo o("Result")("moviedetails")("rating")
	WScript.Echo o("Result")("moviedetails")("runtime")
	WScript.Echo o("Result")("moviedetails")("title")

	For Each i In o("Result")("moviedetails")("studio")
		WScript.Echo i
	Next
	
	For Each i In o("Result")("moviedetails")("writer")
		WScript.Echo i
	Next
	WScript.Echo o("Result")("moviedetails")("year")
	
	
End Sub



Class VbsJson
    'Author: Demon
    'Date: 2012/5/3
    'Website: http://demon.tw
    Private Whitespace, NumberRegex, StringChunk
    Private b, f, r, n, t

    Private Sub Class_Initialize
        Whitespace = " " & vbTab & vbCr & vbLf
        b = ChrW(8)
        f = vbFormFeed
        r = vbCr
        n = vbLf
        t = vbTab

        Set NumberRegex = New RegExp
        NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
        NumberRegex.Global = False
        NumberRegex.MultiLine = True
        NumberRegex.IgnoreCase = True

        Set StringChunk = New RegExp
        StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
        StringChunk.Global = False
        StringChunk.MultiLine = True
        StringChunk.IgnoreCase = True
    End Sub
    
    'Return a JSON string representation of a VBScript data structure
    'Supports the following objects and types
    '+-------------------+---------------+
    '| VBScript          | JSON          |
    '+===================+===============+
    '| Dictionary        | object        |
    '+-------------------+---------------+
    '| Array             | array         |
    '+-------------------+---------------+
    '| String            | string        |
    '+-------------------+---------------+
    '| Number            | number        |
    '+-------------------+---------------+
    '| True              | true          |
    '+-------------------+---------------+
    '| False             | false         |
    '+-------------------+---------------+
    '| Null              | null          |
    '+-------------------+---------------+
    Public Function Encode(ByRef obj)
        Dim buf, i, c, g
        Set buf = CreateObject("Scripting.Dictionary")
        Select Case VarType(obj)
            Case vbNull
                buf.Add buf.Count, "null"
            Case vbBoolean
                If obj Then
                    buf.Add buf.Count, "true"
                Else
                    buf.Add buf.Count, "false"
                End If
            Case vbInteger, vbLong, vbSingle, vbDouble
                buf.Add buf.Count, obj
            Case vbString
                buf.Add buf.Count, """"
                For i = 1 To Len(obj)
                    c = Mid(obj, i, 1)
                    Select Case c
                        Case """" buf.Add buf.Count, "\"""
                        Case "\"  buf.Add buf.Count, "\\"
                        Case "/"  buf.Add buf.Count, "/"
                        Case b    buf.Add buf.Count, "\b"
                        Case f    buf.Add buf.Count, "\f"
                        Case r    buf.Add buf.Count, "\r"
                        Case n    buf.Add buf.Count, "\n"
                        Case t    buf.Add buf.Count, "\t"
                        Case Else
                            If AscW(c) >= 0 And AscW(c) <= 31 Then
                                c = Right("0" & Hex(AscW(c)), 2)
                                buf.Add buf.Count, "\u00" & c
                            Else
                                buf.Add buf.Count, c
                            End If
                    End Select
                Next
                buf.Add buf.Count, """"
            Case vbArray + vbVariant
                g = True
                buf.Add buf.Count, "["
                For Each i In obj
                    If g Then g = False Else buf.Add buf.Count, ","
                    buf.Add buf.Count, Encode(i)
                Next
                buf.Add buf.Count, "]"
            Case vbObject
                If TypeName(obj) = "Dictionary" Then
                    g = True
                    buf.Add buf.Count, "{"
                    For Each i In obj
                        If g Then g = False Else buf.Add buf.Count, ","
                        buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
                    Next
                    buf.Add buf.Count, "}"
                Else
                    Err.Raise 8732,,"None dictionary object"
                End If
            Case Else
                buf.Add buf.Count, """" & CStr(obj) & """"
        End Select
        Encode = Join(buf.Items, "")
    End Function

    'Return the VBScript representation of ``str(``
    'Performs the following translations in decoding
    '+---------------+-------------------+
    '| JSON          | VBScript          |
    '+===============+===================+
    '| object        | Dictionary        |
    '+---------------+-------------------+
    '| array         | Array             |
    '+---------------+-------------------+
    '| string        | String            |
    '+---------------+-------------------+
    '| number        | Double            |
    '+---------------+-------------------+
    '| true          | True              |
    '+---------------+-------------------+
    '| false         | False             |
    '+---------------+-------------------+
    '| null          | Null              |
    '+---------------+-------------------+
    Public Function Decode(ByRef str)
        Dim idx
        idx = SkipWhitespace(str, 1)

        If Mid(str, idx, 1) = "{" Then
            Set Decode = ScanOnce(str, 1)
        Else
            Decode = ScanOnce(str, 1)
        End If
    End Function
    
    Private Function ScanOnce(ByRef str, ByRef idx)
        Dim c, ms

        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "{" Then
            idx = idx + 1
            Set ScanOnce = ParseObject(str, idx)
            Exit Function
        ElseIf c = "[" Then
            idx = idx + 1
            ScanOnce = ParseArray(str, idx)
            Exit Function
        ElseIf c = """" Then
            idx = idx + 1
            ScanOnce = ParseString(str, idx)
            Exit Function
        ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = Null
            Exit Function
        ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = True
            Exit Function
        ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
            idx = idx + 5
            ScanOnce = False
            Exit Function
        End If
        
        Set ms = NumberRegex.Execute(Mid(str, idx))
        If ms.Count = 1 Then
            idx = idx + ms(0).Length
            ScanOnce = CDbl(ms(0))
            Exit Function
        End If
        
        Err.Raise 8732,,"No JSON object could be ScanOnced"
    End Function

    Private Function ParseObject(ByRef str, ByRef idx)
        Dim c, key, value
        Set ParseObject = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)
        
        If c = "}" Then
            Exit Function
        ElseIf c <> """" Then
            Err.Raise 8732,,"Expecting property name"
        End If

        idx = idx + 1
        
        Do
            key = ParseString(str, idx)

            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) <> ":" Then
                Err.Raise 8732,,"Expecting : delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            ParseObject.Add key, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "}" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            c = Mid(str, idx, 1)
            If c <> """" Then
                Err.Raise 8732,,"Expecting property name"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
    End Function
    
    Private Function ParseArray(ByRef str, ByRef idx)
        Dim c, values, value
        Set values = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "]" Then
            ParseArray = values.Items
            Exit Function
        End If

        Do
            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            values.Add values.Count, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "]" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
        ParseArray = values.Items
    End Function
    
    Private Function ParseString(ByRef str, ByRef idx)
        Dim chunks, content, terminator, ms, esc, char
        Set chunks = CreateObject("Scripting.Dictionary")

        Do
            Set ms = StringChunk.Execute(Mid(str, idx))
            If ms.Count = 0 Then
                Err.Raise 8732,,"Unterminated string starting"
            End If
            
            content = ms(0).Submatches(0)
            terminator = ms(0).Submatches(1)
            If Len(content) > 0 Then
                chunks.Add chunks.Count, content
            End If
            
            idx = idx + ms(0).Length
            
            If terminator = """" Then
                Exit Do
            ElseIf terminator <> "\" Then
                Err.Raise 8732,,"Invalid control character"
            End If
            
            esc = Mid(str, idx, 1)

            If esc <> "u" Then
                Select Case esc
                    Case """" char = """"
                    Case "\"  char = "\"
                    Case "/"  char = "/"
                    Case "b"  char = b
                    Case "f"  char = f
                    Case "n"  char = n
                    Case "r"  char = r
                    Case "t"  char = t
                    Case Else Err.Raise 8732,,"Invalid escape"
                End Select
                idx = idx + 1
            Else
                char = ChrW("&H" & Mid(str, idx + 1, 4))
                idx = idx + 5
            End If

            chunks.Add chunks.Count, char
        Loop

        ParseString = Join(chunks.Items, "")
    End Function

    Private Function SkipWhitespace(ByRef str, ByVal idx)
        Do While idx <= Len(str) And _
            InStr(Whitespace, Mid(str, idx, 1)) > 0
            idx = idx + 1
        Loop
        SkipWhitespace = idx
    End Function

End Class


