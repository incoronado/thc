Option Explicit
'On Error Resume Next
Dim SleepVar, Action

SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
Do
	Sleep SleepVar
  	Action = GetPropertyValue ("IR Script.Action")
	If Action <> "Idle" Then	
		SystemCommand(Action)
		SetpropertyValue "IR Script.Action", "Idle"
	End If
	
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
				Case "westinghousetv"
 					Select Case b(1)	
 						Case "power"
 							SetPropertyValue "USBUIRT.Westinghouse Remote", "Power"
 					End Select
 				Case "panasonictv"
 					'IR.System.10.PanasonicTV:On
 					Select Case b(1)	
 						Case "on"
 							SetpropertyValue "USBUIRT.Panasonic Remote", "Power On"
 						Case "off"
 							SetpropertyValue "USBUIRT.Panasonic Remote", "Power Off"
 					End Select
 				Case "sonybluray"
 					'IR.System.10.SonyBluRay:On
 					Select Case b(1)	
 						Case "on"
 							SetpropertyValue "USBUIRT.Sony BD", "Power On"			
 						Case "off"
 							SetpropertyValue "USBUIRT.Sony BD", "Power Off"
 					End Select
 			End Select
	End Select
		
End Sub