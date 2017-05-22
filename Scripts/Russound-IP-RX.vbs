Option Explicit
'On Error Resume Next


Dim Data, SleepVar
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
ReadIPData GetPropertyValue("Russound IP.Received Data")
Sleep SleepVar


Sub ReadIPData(Data)
	Dim Result, line
	listLines = Split(Data, vbCrLf)
	For Each line In listLines
		SePropertyValue "IP Message", line
	Next
End Sub