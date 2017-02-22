Option Explicit
'On Error Resume Next
Dim SleepVar, Action

SleepVar = 5
Do
	Action = GetPropertyValue ("System.Action")
	If Action <> "Idle" Then	
		Call SystemCommand(Action)
		SetpropertyValue "System.Action", "Idle"
	End If
	Sleep SleepVar
Loop


Sub SystemCommand(Action)

End Sub