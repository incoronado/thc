Option Explicit
'On Error Resume Next


Dim Data, SleepVar
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
ReadIPData GetPropertyValue("Russound IP.Received Data")
Sleep SleepVar


Sub ReadIPData(Data)
	If Left(Data, 2) = "N " Then
		SetPropertyValue "Russound IP.IP Message", GetPropertyValue("Russound IP.MessageStr") & Data
		If GetPropertyValue("Russound IP.MessageStr") <> "" Then
			SetPropertyValue "Russound IP.MessageStr", ""
		End if	
	Else
		SetPropertyValue "Russound IP.MessageStr", GetPropertyValue("Russound IP.MessageStr") & Data	
	End if
End Sub