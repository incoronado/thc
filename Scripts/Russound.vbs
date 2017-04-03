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
	Dim a, b, zerobasedzone
	'Class.Source.Priority.Command:Parameters
    'Example:  Russound.Desktop.10.ZonePower:1:on
	a=split(Action,".")
	
    'MsgBox "Ready Command"
    Select Case lcase(a(0))
		Case "russound"
			b=split(a(3),":")
			'Russound.GalaxyTabA1.10.ZonePower:1:on
			Select Case lcase(b(0))
				Case "zonepower"
					Select Case lcase(b(2))
						Case "on"
							zerobasedzone = Right("0" & cStr(cInt(b(1)) - 1),2)
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 F1 23 00 01 00 " & zerobasedzone & " 00 01"
						Case "off"
							zerobasedzone = Right("0" & cStr(cInt(b(1)) - 1),2)
							SerialCommand "F0 00 00 7F 00 " & zerobasedzone & " 70 05 02 02 00 00 F1 23 00 00 00 " & zerobasedzone & " 00 01"
						Case "toggle"

					End Select	
				'Russound.GalaxyTabA1.10.Tune:Up		
				Case "tune"
					Select Case lcase(b(1))
						Case "up"
							SerialCommand "F0 00 00 7F 00 01 70 05 02 02 00 00 69 00 00 00 00 00 01"
						Case "down"
							SerialCommand "F0 00 00 7F 00 01 70 05 02 02 00 00 6A 00 00 00 00 00 01"
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
