Option Explicit
'On Error Resume Next


Dim Data, SleepVar
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
ReadIPData GetPropertyValue("Russound IP.Received Data")
Sleep SleepVar


Sub ReadIPData(Data)
	Dim ListLines, line
	ListLines = Split(Data, vbCrLf)
	For Each line In ListLines
		SetPropertyValue "IP Message", line
	Next
End Sub