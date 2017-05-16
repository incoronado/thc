Option Explicit
On Error Resume Next

Dim Action
Dim SleepVar
Dim objDB, HousebotLocation

HousebotLocation = "C:\Program Files (x86)\housebot\"

'Set su = CreateObject("newObjects.utilctls.StringUtilities")

'-------------------------------------------------------
'- Main: Checks Received Data and Handles Send Data ----
'-------------------------------------------------------
SleepVar = 50

Do
	Action = GetPropertyValue ("UPB Script.Action")
	If Action <> "Idle" Then
		SetpropertyValue "UPB Script.Action", "Idle"
		MessageHandler(Action)
	End If
    Sleep SleepVar
Loop

Sub MessageHandler(Action)
	dim a, b
	Set objDB  = CreateObject("newObjects.sqlite3.dbutf8")
	objDB.Open(HousebotLocation & "config\scripts\lights.mlf")
	objDB.BusyTimeout=1000
	objDB.AutoType = True
	objDB.TypeInfoLevel = 4

	a=split(Action,".") 
    Select Case a(0)
		' Get Class
		Case "UPB"
			b=split(a(3),":")
			select case b(0)
				Case "SelectLight"
					'i.e.,UPB.GalaxyTabA1.10.SelectLight
					SelectLight GetRemoteNumber(a(1))
				Case "PopulateLightList"
					'i.e.,UPB.GalaxyTabA1.10.PopulateLightList
					PopulateLightList "False", GetRemoteNumber(a(1))
				Case "PopulateLightListWithDrilldown"
					'i.e.,UPB.GalaxyTabA1.10.PopulateLightListWithDrillDown
					PopulateLightList "False", GetRemoteNumber(a(1))
				Case "ChangeListContext"
					'i.e., UPB.GalaxyTabA1.10.ChangeListContext:All or Room or Floor
					ChangeListContext b(1), GetRemoteNumber(a(1))
				Case "LightsDoubleClick"			
					'i.e., UPB.GalaxyTabA1.10.LightsDoubleClick 
					LightsDoubleClick GetRemoteNumber(a(1))
				Case "Status"
					'i.e., UPB.GalaxyTabA1.10.RequestUPBStatus:01 
					RequestUPBStatus "DA", Ucase(b(1))
				Case "UPBLightToggle"
					'i.e., UPB.GalaxyTabA1.10.UPBLightToggle:01
					UPBLightToggle Ucase(b(1)), GetRemoteNumber(a(1))
				Case "SelectedUPBLightToggle"
					'i.e., UPB.GalaxyTabA1.10.SelectedUPBLightToggle
					UPBLightToggle ucase(GetPropertyValue(GetRemoteNumber(a(1)) & ".UPB Light ID")), GetRemoteNumber(a(1))
				Case "UPBLightOn"
					'i.e., UPB.GalaxyTabA1.10.UPBLightOn:01
					UPBLightOn Ucase(b(1)), GetRemoteNumber(a(1))
				Case "UPBSceneOn"
					'i.e., UPB.GalaxyTabA1.10.UPBSceneOn:01
					UPBSceneOn Ucase(b(1)), GetRemoteNumber(a(1))	
				Case "UPBSceneOff"
					'i.e., UPB.GalaxyTabA1.10.UPBSceneOff:01
					UPBSceneOff Ucase(b(1)), GetRemoteNumber(a(1))	
					
				Case "SelectedUPBLightOn"
					'i.e., UPB.GalaxyTabA1.10.SelectedUPBLightOn
					UPBLightOn ucase(GetPropertyValue(GetRemoteNumber(a(1)) & ".UPB Light ID")), GetRemoteNumber(a(1))
				Case "UPBLightOff"
					'i.e., UPB.GalaxyTabA1.10.UPBLightOff:01
					UPBLightOff Ucase(b(1)), GetRemoteNumber(a(1))
				Case "SelectedUPBLightOff"
					'i.e., UPB.GalaxyTabA1.10.SelectedUPBLightOn 
					UPBLightOff ucase(GetPropertyValue(GetRemoteNumber(a(1)) & ".UPB Light ID")), GetRemoteNumber(a(1))	
				Case "UPBLightDim"
					'i.e., UPB.GalaxyTabA1.10.UPBLightDim:20
					UPBLightDim ucase(GetPropertyValue(GetRemoteNumber(a(1)) & ".UPB Light ID")) , b(1), GetRemoteNumber(a(1))
				Case "UPBLightIDDim"
					'i.e., UPB.GalaxyTabA1.10.UPBLightIDDim:0D:50
					UPBLightDim b(1), b(2), GetRemoteNumber(a(1))
				Case "SelectedUPBLightDim"
					'i.e., UPB.GalaxyTabA1.10.SelectedUPBLightDim
					UPBLightdim ucase(GetPropertyValue(GetRemoteNumber(a(1)) & ".UPB Light ID")), GetPropertyValue (GetRemoteNumber(a(1)) & ".UPB Dim Level"),GetRemoteNumber(a(1))
				Case "PollAllUPBLights"
					'i.e., UPB.GalaxyTabA1.10.PollAllUPBLights
					PollAllUPBLights
			End Select
	
		Case "upblight"
			select case a(1)
				Case "status"
					RequestUPBStatus "DA", Ucase(Mid(a(2),5))
				
				Case "toggle-galaxyprotab1"
				    UPBLightToggle ucase(GetPropertyValue("GalaxyProTab1 LightControl.UPB Light ID"))
				Case "toggle-chalkboard"
				    UPBLightToggle ucase(GetPropertyValue("Chalkboard LightControl.UPB Light ID"))	
					
				Case "toggle"
					UPBLightToggle Ucase(Mid(a(2),5))
				
				Case "on-galaxyprotab1"
					UPBLightOn	ucase(GetPropertyValue("GalaxyProTab1 LightControl.UPB Light ID"))
				Case "on-chalkboard"
					UPBLightOn	ucase(GetPropertyValue("Chalkboard LightControl.UPB Light ID"))	
				
				Case "on"
					UPBLightOn Ucase(Mid(a(2),5))
				
				Case "off-galaxyprotab1"
					UPBLightOff	ucase(GetPropertyValue("GalaxyProTab1 LightControl.UPB Light ID"))
				Case "off-chalkboard"
					UPBLightOff	ucase(GetPropertyValue("Chalkboard LightControl.UPB Light ID"))
				
				Case "off"
					UPBLightOff Ucase(Mid(a(2),5))
				
				Case "dim-galaxyprotab1"
					UPBLightdim ucase(GetPropertyValue("GalaxyProTab1 LightControl.UPB Light ID")), GetPropertyValue ("GalaxyProTab1 LightControl.Dim Level")	
				Case "dim-chalkboard"
					UPBLightdim ucase(GetPropertyValue("Chalkboard LightControl.UPB Light ID")), GetPropertyValue ("Chalkboard LightControl.Dim Level")		
					
					
				Case "dim"
					UPBLightdim ucase(Mid(a(2),5)), GetPropertyValue ("UPB-" + UCase(Mid(a(2),5)) + ".Dim Level")
				Case "poll"
					PollAllUPBLights
			
			End Select	
		Case "populate light list"
			PopulateLightList "False"	
		Case "populate light list with drilldown"
			PopulateLightList "True"
		Case "select light"
			SelectLight
		Case "change list context floor"
			ChangeListContext "Floor"	
		Case "change list context room"	
		    ChangeListContext "Room", ""	
		Case "change list context all"		
		    ChangeListContext "All"
		Case "lights double-click"				
		    LightsDoubleClick
	End Select
	objDB.Close
    Set objDB = Nothing
	   
End Sub


'PU0800DAFF0686642F
'PU0800DAFF12866423

Function PollAllUPBLights
	Dim LightCount, i
	LightCount = CountUPBLights()
	For i=1 to LightCount
		RequestUPBStatus "DA", right("0" + hex(i),2)
		Sleep 1000
	Next
End Function

Sub SelectLight(Remote)
	'SetPropertyValue "UPB Script.Debug", Remote & ".Room Context"
	Dim LightList, SqlStr, r, Row
	LightList = ""
	SqlStr = "SELECT lights.Name, floor.Floor, room.Room, lights.UPB_ID FROM lights left JOIN room ON room.id=lights.RoomID left join floor on floor.id=room.FloorID where lights.id = " & GetPropertyValue(Remote & ".Selected Light") & " order by Floor ASC, Room ASC, Name ASC"
	Set r = objDB.Execute(SqlStr) 
	For Row = 1 To r.count
	  SetPropertyValue Remote & ".Room Context",  r(Row)("Room")
	  SetPropertyValue Remote & ".UPB Light Name",  r(Row)("Name")
	  SetPropertyValue Remote & ".UPB Dim Level",  GetpropertyValue("UPB-" & r(Row)("UPB_ID") & ".Dim Level")
	  SetPropertyValue Remote & ".UPB Light State",  GetpropertyValue("UPB-" & r(Row)("UPB_ID") & ".Power State")
      SetPropertyValue Remote & ".UPB Light ID",  r(Row)("UPB_ID")
	Next   
End Sub

Sub ChangeListContext(ListContext, Remote)
   SetPropertyValue Remote & ".List Context",  ListContext
   PopulateLightList "False", Remote
End Sub

Sub LightsDoubleClick(Remote)
   If GetPropertyValue(Remote & ".List Context") = "Floor" Then
      SetPropertyValue Remote & ".List Context", "Room"
      PopulateLightList "True", Remote
   ElseIf GetPropertyValue(Remote & ".List Context") = "Room" Then
      SetPropertyValue Remote & ".List Context", "Light"   
      PopulateLightList "True", Remote
   End if
End Sub

Sub PopulateLightList(drilldown,Remote) 
	Dim FloorList, RoomList, LightList, SqlStr, r, Row, WhereStr, Col1, Col2
	LightList = ""
	WhereStr = ""
	
	If drilldown = "True" Then
		If GetPropertyValue(Remote & ".List Context") = "Room" Then
			SqlStr = "SELECT * from room where FloorID = " & CInt(GetPropertyValue(Remote & ".Lights Double-Click Value"))
			Col1 = "id"
			Col2 = "Room"
			SetPropertyValue Remote & ".Selected Light", 1
		Elseif GetPropertyValue(Remote & ".List Context") = "Light" Then
			WhereStr = " where lights.RoomID = " & CInt(GetPropertyValue(Remote & ".Lights Double-Click Value"))
			SetPropertyValue Remote & ".Selected Light", 1
		    SqlStr = "SELECT lights.ID, lights.Name, lights.RoomID, room.FloorID, lights.UPB_ID, floor.Floor, room.Room from lights left JOIN room ON room.id=lights.RoomID left join floor on floor.id=room.FloorID" & WhereStr & " order by floor.Floor asc, room.Room asc, lights.Name asc"	
			Col1 = "id"
		    Col2 = "Name"
		End if
		

		If GetPropertyValue(Remote & ".List Context") = "Room" Then	
			SetPropertyValue Remote & ".Lights Dialog Title", GetPropertyValue(Remote & ".Floor Context") & "->" & GetPropertyValue(Remote & ".Room Context") 
		ElseIf GetPropertyValue(Remote & ".List Context") = "Floor" Then		
			SetPropertyValue Remote & ".Lights Dialog Title", GetPropertyValue(Remote & ".Floor Context")
		ElseIf GetPropertyValue(Remote & ".List Context") = "All" Then		
			SetPropertyValue Remote & ".Lights Dialog Title", "All"
		End if	
		
	Else
		If GetPropertyValue(Remote & ".List Context") = "Floor" Then
			SqlStr = "SELECT * from floor order by Floor"
			Col1 = "id"
			Col2 = "Floor"
			SetPropertyValue Remote & ".Selected Light", 1
		ElseIf GetPropertyValue(Remote & ".List Context") = "Room" Then	
			SqlStr = "SELECT * from room order by Room"
			SetPropertyValue Remote & ".Selected Light", 7
			Col1 = "id"
			Col2 = "Room"
		ElseIf GetPropertyValue(Remote & ".List Context") = "Light" Then 
		    SqlStr = "SELECT lights.ID, lights.Name, lights.RoomID, room.FloorID, lights.UPB_ID, floor.Floor, room.Room from lights left JOIN room ON room.id=lights.RoomID left join floor on floor.id=room.FloorID" & WhereStr & " order by floor.Floor asc, room.Room asc, lights.Name asc"
			Col1 = "id"
			Col2 = "Name"
			SetPropertyValue Remote & ".Selected Light", 16
		ElseIf GetPropertyValue(Remote & ".List Context") = "All" Then 
		    SqlStr = "SELECT lights.ID, lights.Name, lights.RoomID, room.FloorID, lights.UPB_ID, floor.Floor, room.Room from lights left JOIN room ON room.id=lights.RoomID left join floor on floor.id=room.FloorID order by floor.Floor asc, room.Room asc, lights.Name asc"
			Col1 = "id"
			Col2 = "Name"
			SetPropertyValue Remote & ".Selected Light", 16
		End if
        If GetPropertyValue(Remote & ".List Context") = "Room" Then
			SetPropertyValue Remote & ".Lights Dialog Title", "Rooms"
		ElseIf GetPropertyValue(Remote & ".List Context") = "Floor" Then
			SetPropertyValue Remote & ".Lights Dialog Title", "Floors"
		ElseIf GetPropertyValue(Remote & ".List Context") = "All" Then
			SetPropertyValue Remote & ".Lights Dialog Title", "All"
		End if	
	End If	

	Set r = objDB.Execute(SqlStr)
	If r.Count = 0 Then
	   LightList     = " " & vbTAB & "No Light List Found" & vbLF
	End if 
	For Row = 1 To r.count
		If r(Row)(Col1) = CInt(GetPropertyValue(Remote & ".Selected Light")) Then
			LightList = LightList & "*S-" & r(Row)(Col1) & vbTAB & r(Row)(Col2) & vbLF 
		Else
			LightList = LightList & r(Row)(Col1) & vbTAB & r(Row)(Col2) & vbLF 
		End if
	Next
	Set r = Nothing
	SetPropertyValue Remote & ".Light List", Left(LightList,Len(LightList)-1)
	
End Sub

Function CountUPBLights
	Dim LightCount,NoMoreLights
	NoMoreLights = 0
	LightCount = 0
	'  Warning -- assumes light IDs are sequential and there are no ID gaps
	Do Until NoMoreLights=1
		If GetPropertyValue ("UPB-" + right("0" + hex(LightCount+1),2) + ".Power State") <> "* error *" Then
			LightCount=LightCount + 1
		Else
			NoMoreLights = 1
		End if
	Loop
	CountUPBLights = LightCount
End Function

Function UPBLightOn(UPBLightID, Remote) 
		SendUPBCommand "DA", UPBLightID, "64"
		SetpropertyValue "UPB-" & UPBLightID & ".Power State", "On"
		SetpropertyValue "UPB-" & UPBLightID & ".Dim Level", "100"
		UpdateRemoteSettings UPBLightID, Remote
End Function

Function UPBSceneOn(UPBLightID, Remote) 
		SendSceneCommandOn "DA", UPBLightID
		'SetpropertyValue "UPB-" & UPBLightID & ".Power State", "On"
		'SetpropertyValue "UPB-" & UPBLightID & ".Dim Level", "100"
		'UpdateRemoteSettings UPBLightID, Remote
End Function

Function UPBSceneOff(UPBLightID, Remote) 
		SendSceneCommandOff "DA", UPBLightID
		'SetpropertyValue "UPB-" & UPBLightID & ".Power State", "On"
		'SetpropertyValue "UPB-" & UPBLightID & ".Dim Level", "100"
		'UpdateRemoteSettings UPBLightID, Remote
End Function


Function UPBLightOff(UPBLightID, Remote) 
		SendUPBCommand "DA", UPBLightID, "00"
		SetpropertyValue "UPB-" & UPBLightID & ".Power State", "Off"
		SetpropertyValue "UPB-" & UPBLightID & ".Dim Level", "0"
		UpdateRemoteSettings UPBLightID, Remote	
End Function

Function UPBLightDim(UPBLightID, DimLevel, Remote)
		If DimLevel = 0 then
			SetpropertyValue "UPB-" & UPBLightID & ".Power State", "Off"	
		Else 
			SetpropertyValue "UPB-" & UPBLightID & ".Power State", "On"	
		End if		
		SendUPBCommand "DA", UPBLightID, Right("0" & hex(DimLevel),2)
		SetpropertyValue "UPB-" & UPBLightID & ".Dim Level", DimLevel
		UpdateRemoteSettings UPBLightID, Remote 	
End Function


Function UPBLightToggle(UPBLightID, Remote) 
	If GetpropertyValue("UPB-" & UPBLightID & ".Power State") = "Off" then
		SendUPBCommand "DA", UPBLightID, "64"
		SetpropertyValue "UPB-" & UPBLightID & ".Power State", "On"
		SetpropertyValue "UPB-" & UPBLightID & ".Dim Level", "100"
	Else 	
		SendUPBCommand "DA", UPBLightID, "00"
		SetpropertyValue "UPB-" & UPBLightID & ".Power State", "Off"
		SetpropertyValue "UPB-" & UPBLightID & ".Dim Level", "0"
	End If
    UpdateRemoteSettings UPBLightID, Remote
End Function

Function UpdateRemoteSettings (UPBLightID, Remote) 
	Dim RemoteList, i
	RemoteList = split(GetPropertyValue("System.HBRemoteList"),",")
	
	For Each i in RemoteList
		If GetPropertyValue("Remote-" & CStr(i) & ".UPB Light ID") = UPBLightID Then
			SetPropertyValue "Remote-" & CStr(i) & ".UPB Light State", GetpropertyValue("UPB-" & UPBLightID & ".Power State")
			SetPropertyValue "Remote-" & CStr(i) & ".UPB Dim Level", GetpropertyValue("UPB-" & UPBLightID & ".Dim Level")
		End if
	
	Next

	

End Function

Function RequestUPBStatus(NetworkID, DeviceID)
	SendSerialCommand "0700" & NetworkID & DeviceID & "FF30"
	'SetpropertyValue "UPB Script.Debug" , "0710" & NetworkID & DeviceID & "FF30"
End Function

Function SendUPBCommand(NetworkID, DeviceID, LightLevel)
	SendSerialCommand "0900" & NetworkID & DeviceID & "FF23" & LightLevel & "00"
End Function

Function SendSceneCommandOn(NetworkID, DeviceID)
	SendSerialCommand "8700" & NetworkID & DeviceID & "FF20"
End Function

Function SendSceneCommandOff(NetworkID, DeviceID)
	SendSerialCommand "8700" & NetworkID & DeviceID & "FF21"
End Function
Function SendSerialCommand (data)
   	SetpropertyValue "UPB PIM.Serial Command", chr(20) & data & UPB_Checksum(data)
End Function 

Function UPB_Checksum(hexstr)
	Dim ComputedChecksum, i
	'Calculate Checksum
	ComputedChecksum=0
	    For i = 1 To len(hexstr) Step 2
			ComputedChecksum = ComputedChecksum + CLng("&h" & Mid(hexstr,i,2))
		Next
	UPB_Checksum=right(TwosComplement(cStr(hex(ComputedChecksum))),2)
End Function

Function TwosComplement(value)
    TwosComplement=hex(CLng("&h" & OnesComplement(value))+1)
End Function

Function OnesComplement(value)
	Dim OnesCompStr, i
    OnesCompStr = ""
	for i = 1 to len(value)
	 Select Case mid(value,i,1) 
	' Configuration Command
	   Case "0"
	     OnesCompStr = OnesCompStr + "F"
		Case "1"
	     OnesCompStr = OnesCompStr + "E" 
		Case "2"
	     OnesCompStr = OnesCompStr + "D" 
		Case "3"
	     OnesCompStr = OnesCompStr + "C" 
		Case "4"
	     OnesCompStr = OnesCompStr + "B" 
		Case "5"
	     OnesCompStr = OnesCompStr + "A" 
		Case "6"
	     OnesCompStr = OnesCompStr + "9" 
		Case "7"
	     OnesCompStr = OnesCompStr + "8"
		Case "8"
	     OnesCompStr = OnesCompStr + "7"
		Case "9"
	     OnesCompStr = OnesCompStr + "6"
		Case "A"
	     OnesCompStr = OnesCompStr + "5"
		Case "B"
	     OnesCompStr = OnesCompStr + "4"
		Case "C"
	     OnesCompStr = OnesCompStr + "3"
		Case "D"
	     OnesCompStr = OnesCompStr + "2"
		Case "E"
	     OnesCompStr = OnesCompStr + "1" 
		Case "F"
	     OnesCompStr = OnesCompStr + "0"
	   End Select 
	Next
	OnesComplement = OnesCompStr
End Function

'WScript.Echo TwoComplement8Bits("&hf8")
Function Hex2Dec(hex_value)
	Hex2Dec=CLng("&h" & hex_value)
End Function

Function MSTimer()
	MSTimer=Timer*1000 
End Function

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
