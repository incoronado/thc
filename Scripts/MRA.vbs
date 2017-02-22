Option Explicit
'On Error Resume Next

Dim Action
Dim SleepVar

'-------------------------------------------------------
'- Main: Checks Received Data and Handles Send Data ----
'-------------------------------------------------------

SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))

Do
	Action = GetPropertyValue ("Multiroom Audio Script.Action")
	If Action <> "Idle" Then
		SetpropertyValue "Multiroom Audio Script.Action", "Idle"
		MessageHandler(Action)

	End If
	Sleep SleepVar
	
Loop

 
Sub MessageHandler(Action)
	Dim a, b, SendStr
	'Class.Source.Priority.Command:Parameters
    'Example:  MRA.RS232.10.ProcessSerialMessage
	a=split(Action,".")
	Select Case a(0)
		Case "MRA"
			b=split(a(3),":")
			if lcase(b(0)) <> "poll" Then
				'Turn Off polling if any other command than poll is sent
				SetpropertyValue "Block MRA Poll.Running", "Yes"
				SetModeState "Block MRA Poll", "Inactive"	
			End if
			Select Case lcase(b(0))	
				Case "poll"
					'MRA.GalaxyTabA1.10.poll:2
					SendStr = "?1" & b(1)
					SendSerialCommand SendStr
				Case "on"
					'MRA.GalaxyTabA1.10.on:2
					SendStr = "<1" & b(1) & "PR01"
					SendSerialCommand SendStr
				Case "off"
					'MRA.GalaxyTabA1.10.off:2
					SendStr = "<1" & b(1) & "PR00"
					SendSerialCommand SendStr
				Case "vol"
					'MRA.GalaxyTabA1.10.vol:2:15
					SendStr = "<1" & b(1) & "VO" & right("0" + CStr(b(2)),2)
					SendSerialCommand SendStr
				Case "source"
					'MRA.GalaxyTabA1.10.source:2:5
					SendStr = "<1" & b(1) & "CH" & right("0" + CStr(b(2)),2)
					SendSerialCommand SendStr
				Case "mute"	
					'MRA.GalaxyTabA1.10.mute:2:on
					Select Case lcase(b(2))
						Case "on"
							SendStr = "<1" & b(1) & "MU01"
							SendSerialCommand SendStr
						Case "off"
							SendStr = "<1" & b(1) & "MU00"
							SendSerialCommand SendStr
					End Select	
					
				
				
			End Select
	End Select
	
End Sub

Sub SendSerialCommand(Data)
	SetpropertyValue "Multiroom Audio Amplifier.MRA Command", Data
	sleep 5
End Sub

