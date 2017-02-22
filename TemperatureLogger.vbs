Option Explicit
On Error Resume Next
Dim Action, Status, SleepVar

SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))

Do
	Action = GetPropertyValue ("Temperature.Action")
	If Action <> "Idle" Then	
		SetpropertyValue "Temperature.Action", "Idle"
		Message_Handler(Action)
	End If
	Sleep SleepVar
Loop


Sub Message_Handler(Message)
Dim a, b
	a=split(Message,".")
	Select Case a(0)
		'Temperature Class Message Received
		'Temperature.GalaxyTabA1.ReadTemps
		Case "Temperature"
		
			b=split(a(3),":")
			'System.GalaxyTabA1.10.SelectZone:1
			Select Case b(0)
				Case "ProcessSerialMessage"
					ProcessSerialMessage
			End Select 
	End Select
 End Sub
 
 
 Sub ProcessSerialMessage
	Dim Temps
	Temps = split(GetPropertyValue ("Temperature Logger.Received Data"), ", ")
	SetPropertyValue "Temperature.Cabinet Temperature",  Left(Mid(Temps(3),4), Len(Mid(Temps(3),4)) - 1) & chr(176)
 End Sub

