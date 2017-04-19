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
	Action = GetPropertyValue ("Multiroom Audio Settings.Action")
	If Action <> "Idle" Then
		MessageHandler(Action)
		SetpropertyValue "Multiroom Audio Settings.Action", "Idle"
	End If
	Sleep SleepVar
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
				Case "zonepower"
					zerobasedzone = Zone2ID(Right("0" & cStr(cInt(b(1)) - 1),2))
					Select Case lcase(b(2))
						Case "on"		
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 F1 23 00 01 00 " & zerobasedzone & " 00 01"
						Case "off"
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 F1 23 00 00 00 " & zerobasedzone & " 00 01"
						Case "toggle"
					End Select	
				'Russound.GalaxyTabA1.10.Keypad:1:Previous		
				Case "keypad"
					zerobasedzone = Right("0" & cStr(cInt(b(1)) - 1),2)
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
					'Russound.GalaxyTabA1.10.ZoneSource:1:1
					zerobasedzone = Right("0" & cStr(cInt(b(1)) - 1),2)
					zerobasedsource = Right("0" & cStr(cInt(b(2)) - 1),2)
					SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 00 00 00 F1 3E 00 00 00 " & zerobasedsource & " 00 01"
				Case "volume"
					'Russound.GalaxyTabA1.10.Volume:1:up
					zerobasedzone = Right("0" & cStr(cInt(b(1)) - 1),2)
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

	SetPropertyValue "Multiroom Audio Settings.Debug 2", hexstr & ComputeRNETChecksum(hexstr) & "F7"
	SetPropertyValue "Multiroom Audio Amplifier.MRA Command", Replace(hexstr," ","",1,-1) & ComputeRNETChecksum(hexstr) & "F7"

End Sub

Function SourceName2ID(SourceName) 
	Dim SourceCount, NoMoreSources, SourceNo
	SourceNo=0
	NoMoreSources = 0
	SourceCount = 0
	Do Until NoMoreEventSources = 1
		If GetPropertyValue("Multiroom Audio Settings.Zone " + CStr(SourceCount+1) + " Name") <> "* error *" Then
		   If GetPropertyValue("Multiroom Audio Settings.Zone " + CStr(EventHandlerCount+1) + " Name") = SourceName Then
		   		SourceNo = SourceCount
		   End if
		   SourceCount=SourceCount + 1  
		Else
			NoMoreSources = 1
		End if
	Loop
	SourceName2ID = SourceNo
End Sub

Function ZoneName2ID(ZoneName) 
	Dim ZoneCount, NoMoreZones, ZoneNo
	ZoneNo=0
	NoMoreZones = 0
	ZoneCount = 0
	Do Until NoMoreZones = 1
		If GetPropertyValue("Multiroom Audio Settings.Zone " + CStr(ZoneCount+1) + " Name") <> "* error *" Then
		   If GetPropertyValue("Multiroom Audio Settings.Zone " + CStr(EventHandlerCount+1) + " Name") = ZoneName Then
		   		ZoneNo = ZoneCount
		   End if
		   ZoneCount=ZoneCount + 1  
		Else
			NoMoreZones = 1
		End if
	Loop
	ZoneName2ID = ZoneNo
End Sub



