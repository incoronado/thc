Option Explicit
'On Error Resume Next
Dim SleepVar, Action

SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
Do
  	Action = GetPropertyValue ("IR Script.Action")
	If Action <> "Idle" Then	
		SetpropertyValue "IR Script.Action", "Idle"
		SystemCommand(Action)
	End If
	Sleep SleepVar
Loop

Sub SystemCommand(Action)
	Dim a, b
	a=split(lcase(Action),".")
	
	Select Case a(0)
		' IR Class Message Received
		' IR.System.10.WestingHouseTV:Power
		Case "ir"
 			b=split(a(3),":")
 			Select Case b(0)
				Case "westingHousetv"
 					Select Case b(1)	
 						Case "power"
 							SetPropertyValue "USBUIRT.Westinghouse Remote", "Power"
 					End Select
 			End Select

	End Select
		
End Sub