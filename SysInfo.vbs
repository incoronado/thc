Option Explicit
'On Error Resume Next

Dim Action, SleepVar

SleepVar = 5

Do
	Sleep SleepVar
		
	Action = GetPropertyValue ("SysInfo.Action")
	If Action <> "Idle" Then
		SetpropertyValue "SysInfo.Action", "Idle"
		Sleep SleepVar
		If Action <> "" Then
			Call ProcessMessage(Action)
		End If
	End If
Loop


Sub ProcessMessage(Action)
    'MsgBox "Ready Command"
    Select Case Action 
	' Configuration Command
	   Case "GetSysInfo"
	     GetSysInfo
	   End Select
End Sub


Sub GetSysInfo
	Dim strComputer, objWMIService, colItems, objItem 
	strComputer = "."

	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\OpenHardwareMonitor")
	Set colItems = objWMIService.ExecQuery("SELECT value FROM Sensor WHERE Name = 'CPU Core #1' AND SensorType = 'Temperature'")
	For Each objItem In colItems
	   SetpropertyValue "SysInfo.SysInfo-CPU Core 1 Temp", objItem.value
	Next
	
	Set colItems = objWMIService.ExecQuery("SELECT value FROM Sensor WHERE Name = 'CPU Core #2' AND SensorType = 'Temperature'")
	For Each objItem In colItems
	   SetpropertyValue "SysInfo.SysInfo-CPU Core 2 Temp", objItem.value
	Next
	
	Set colItems = objWMIService.ExecQuery("SELECT value FROM Sensor WHERE Name = 'CPU Package' AND SensorType = 'Temperature'")
	For Each objItem In colItems
	   SetpropertyValue "SysInfo.SysInfo-CPU Package Temp", objItem.value
	Next
	

End Sub
