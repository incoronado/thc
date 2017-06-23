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
	Dim a, b, i, zerobasedzone, zerobasedsource
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
					SendIPCommand("WATCH SYSTEM ON")
					SendIPCommand("WATCH C[1].Z[1] ON")
					SendIPCommand("WATCH C[1].Z[2] ON")
					SendIPCommand("WATCH C[1].Z[8] ON")
					SendIPCommand("WATCH C[1].Z[4] ON")
					SendIPCommand("WATCH S[1] ON")
					SendIPCommand("WATCH S[2] ON")
					SendIPCommand("WATCH S[3] ON")
					SendIPCommand("WATCH S[4] ON")
					SendIPCommand("WATCH S[5] ON")
					SendIPCommand("WATCH S[6] ON")
					SendIPCommand("WATCH S[7] ON")
					SendIPCommand("WATCH S[8] ON")
				'Russound.GalaxyTabA1.10.ZonePower:Master Bedroom:On			
				Case "zonepower"
					Select Case lcase(b(2))
						Case "on"
						    SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!ZoneOn"
						    SetPropertyValue "Multiroom Audio Settings.Zone " & cStr(ZoneName2ID(b(1))) & " Power", "On"
						Case "off"
							SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!ZoneOff"
							SetPropertyValue "Multiroom Audio Settings.Zone " & cStr(ZoneName2ID(b(1))) & " Power", "Off"
						Case "toggle"
					End Select	
				'Russound.GalaxyTabA1.10.Keypad:Patio:Previous
				Case "keypad"
					SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!KeyRelease " & b(2)
				'Russound.GalaxyTabA1.10.ZoneSource:Patio:Cable TV
				Case "zonesource"
					SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!SelectSource " & cStr(SourceName2ID(b(2)))
				Case "volume"
					'Russound.GalaxyTabA1.10.Volume:Master Bedroom:up
					Select Case lcase(b(2))
						Case "up"
							SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!KeyPress VolumeUp"
						Case "down"
							SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!KeyPress VolumeDown"
					End Select
				Case "volumeto"
					'Russound.GalaxyTabA1.10.VolumeTo:Master Bedroom:15					
					SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!KeyPress Volume " & b(2)
				'Russound.GalaxyTabA1.10.SendDigit:Master Bedroom:1	
				Case "senddigit"
					SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!KeyRelease Digit" & ConvertDigit(b(2))
				'Russound.GalaxyTabA1.10.SendDigits:Master Bedroom:1007
				Case "senddigits"
					For i = 1 to len(b(2))
				   		SetPropertyValue "Russound IP.Send Data", "EVENT C[1].Z[" & cStr(ZoneName2ID(b(1))) & "]!KeyRelease Digit" & ConvertDigit(Mid(b(2),i,1))
				   		Sleep 150
					Next
			End Select	
	End Select
End Sub		

Sub SendIPCommand(Command)
    Dim StartTime
    StartTime = Timer() 
	SetPropertyValue "Russound IP.Command Successful", 0


	Do Until GetPropertyValue("Russound IP.Ready To Send") = 1
	  If (Timer() - StartTime) >= 3 Then
	    SetPropertyValue "Russound IP.Trace Errors", "IP Channel is Busy"
	    Exit Do
	  End If  
	Loop

    StartTime = Timer() 
	SetPropertyValue "Russound IP.Send Data", Command
	Do Until GetPropertyValue("Russound IP.Command Successful") = 1
	  If (Timer() - StartTime) >= 3 Then
	    SetPropertyValue "Russound IP.Trace Errors", "No Acknowlege Received-" & Command
	    Exit Do
	  End If  
	Loop
	'SetPropertyValue "Russound IP.Command Successful", 1
End Sub




Function ConvertDigit (MyDigit)
   Select Case CInt(MyDigit)
      Case 1: ConvertDigit = "One"
      Case 2: ConvertDigit = "Two"
      Case 3: ConvertDigit = "Three"
      Case 4: ConvertDigit = "Four"
      Case 5: ConvertDigit = "Five"
      Case 6: ConvertDigit = "Six"
      Case 7: ConvertDigit = "Seven"
      Case 8: ConvertDigit = "Eight"
      Case 9: ConvertDigit = "Nine"
      Case 0: ConvertDigit = "Zero"
      Case Else: ConvertDigit = ""
   End Select
End Function

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
