' Russound
' This script Supports Russound RNET Protocol
' ©2017 Mike Larson created (4/1/2017)
' 
' Requirements:
' RS232 Port COnfigured @ 19200,N,8,1 No Flow Control
'
' Limitations:
' 
' 
Option Explicit
'On Error Resume Next

Dim Action
Dim SleepVar

SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))

Do
	Sleep SleepVar
	Action = GetPropertyValue ("Multiroom Audio Settings.Action")
	If Action <> "Idle" Then
		MessageHandler(Action)
		SetpropertyValue "Multiroom Audio Settings.Action", "Idle"
	End If
Loop

Sub MessageHandler(Action)
	Dim a, b, zerobasedzone, zerobasedsource
	'Class.Source.Priority.Command:Parameters
    'Example:  Russound.Desktop.10.ZonePower:Master Bedroom:on
	a=split(Action,".")
	
    'MsgBox "Ready Command"
    Select Case lcase(a(0))
		Case "russound"
			b=split(a(3),":")
			'Russound.GalaxyTabA1.10.ZonePower:1:on
			Select Case lcase(b(0))
				Case "initip"
					'Russound.GalaxyTabA1.10.InitIP
					SetpropertyValue "Russound IP.Send Data", "WATCH SYSTEM ON"
					Sleep 250
					SetPropertyValue "Russound IP.Send Data", "WATCH C[1].Z[1] ON" 
					Sleep 2000
					SetPropertyValue "Russound IP.Send Data", "WATCH C[1].Z[2] ON" 
					Sleep 2000
					SetPropertyValue "Russound IP.Send Data", "WATCH C[1].Z[8] ON"
					Sleep 2000
					SetPropertyValue "Russound IP.Send Data", "WATCH S[1] ON"
					Sleep 250
					SetPropertyValue "Russound IP.Send Data", "WATCH S[2] ON"
					Sleep 250
					SetPropertyValue "Russound IP.Send Data", "WATCH S[3] ON"
					Sleep 250
					SetPropertyValue "Russound IP.Send Data", "WATCH S[4] ON" 
					Sleep 250
					SetPropertyValue "Russound IP.Send Data", "WATCH S[5] ON"
					Sleep 250
					SetPropertyValue "Russound IP.Send Data", "WATCH S[6] ON"
					Sleep 250
					SetPropertyValue "Russound IP.Send Data", "WATCH S[7] ON"
					Sleep 250
					SetPropertyValue "Russound IP.Send Data", "WATCH S[8] ON"
					Sleep 250
							
				Case "zonepower"
					zerobasedzone = Right("0" & cStr(cInt(ZoneName2ID(b(1))) - 1),2)
					Select Case lcase(b(2))
						Case "on"
						    SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!ZoneOn"
							'SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 F1 23 00 01 00 " & zerobasedzone & " 00 01"
						Case "off"
							SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!ZoneOff"
							'SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 F1 23 00 00 00 " & zerobasedzone & " 00 01"
						Case "toggle"
					End Select	
				'Russound.GalaxyTabA1.10.Keypad:Patio:Previous		
				Case "keypad"
					zerobasedzone = Right("0" & cStr(cInt(ZoneName2ID(b(1))) - 1),2)
					Select Case lcase(b(2))
						Case "previous"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 67 00 00 00 00 00 01"
						Case "next"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 68 00 00 00 00 00 01"
						Case "plus"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 69 00 00 00 00 00 01"
						Case "minus"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 6A 00 00 00 00 00 01"
						Case "play"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 73 00 00 00 00 00 01"
						Case "stop"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 6D 00 00 00 00 00 01"
						Case "pause"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 6E 00 00 00 00 00 01"
						Case "F1"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 6F 00 00 00 00 00 01"
						Case "F2"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 70 00 00 00 00 00 01"
					End Select
				Case "zonesource"
					'Russound.GalaxyTabA1.10.ZoneSource:Patio:Cable TV
					zerobasedzone = Right("0" & cStr(cInt(ZoneName2ID(b(1))) - 1),2)
					zerobasedsource = Right("0" & cStr(cInt(SourceName2ID(b(2))) - 1),2)
					SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 00 00 00 F1 3E 00 00 00 " & zerobasedsource & " 00 01"
				Case "volume"
					'Russound.GalaxyTabA1.10.Volume:Master Bedroom:up
					zerobasedzone = Right("0" & cStr(cInt(ZoneName2ID(b(1))) - 1),2)
					Select Case lcase(b(2))
						Case "up"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 7F 00 00 00 00 00 01"
						Case "down"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 F1 7F 00 00 00 00 00 01"
					End Select

			End Select	
	End Select
End Sub		


Function ComputeRNETChecksum(hexstr)
	
	Dim HexBytes, ComputedChecksum, i

	HexBytes=split(hexstr," ")
	For i = 0 To ubound(HexBytes)
		'Add the HEX value of every byte in the message that precedes the Checksum in decimal
		ComputedChecksum = ComputedChecksum + CLng("&h" & HexBytes(i))
	Next
	'Count the number of bytes which precede the Checksum still maintaining decimal
	ComputedChecksum = ComputedChecksum + ubound(HexBytes) + 1
	' This value is then AND-ed with the HEX value 0x007F or 127 Dec
	ComputedChecksum = ComputedChecksum And 127
	ComputeRNETChecksum=cStr(hex(ComputedChecksum))
End Function


Sub SerialCommand(hexstr)

	'SetPropertyValue "Multiroom Audio Settings.Debug 2", hexstr & ComputeRNETChecksum(hexstr) & "F7"
	SetPropertyValue "Multiroom Audio Amplifier.MRA Command", Replace(hexstr," ","",1,-1) & ComputeRNETChecksum(hexstr) & "F7"

End Sub

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
	'SetPropertyValue "Multiroom Audio Settings.Debug 2", ZoneNo
	ZoneName2ID = ZoneNo
End Function



