Option Explicit
'On Error Resume Next


Dim Data, SleepVar
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
ReadIPData GetPropertyValue("Russound IP.Received Data")
Sleep SleepVar


Sub ReadIPData(Data)
	result = Instr(Data, vbCRLF)
	If Result <> 0
		SetPropertyValue "Russound IP.Debug", Result
	End if
End Sub