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
	Dim a, b
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
							SerialCommand "F000007F0001700502020000F123000100010001"
						Case "off"

						Case "toggle"

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
	SetPropertyValue "MRA Command.MRA Command", hexstr & ComputeRNETChecksum(hexstr) & "F7"

End Sub
