Option Explicit
'On Error Resume Next


Dim Data, SleepVar
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
ReadIPData GetPropertyValue("Russound IP.Received Data")
Sleep SleepVar


Sub ReadIPData(Data)
	SetPropertyValue "Russound IP.MessageStr", GetPropertyValue("Russound IP.MessageStr") & Data
	If Instr(Data, vbCRLF) = 0 Then
		SetPropertyValue "Russound IP.MessageStr", ""
	End if
End Sub