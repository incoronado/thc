Option Explicit
'On Error Resume Next


Dim Data, SleepVar
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
ReadIPData GetPropertyValue("Russound IP.Received Data")
Sleep SleepVar


Sub ReadIPData(Data)
	Dim ListLines, line, keyvalue, key, value, ResponseType,keydata, command, Zstr
	ListLines = Split(Data, vbCrLf)
	For Each line In ListLines
		If line <> "" Then
			SetPropertyValue "Russound IP.IP Message", line
			keyvalue=split(line, "=")
			ResponseType = Left(keyvalue[0],1)
			key = Mid(keyvalue[0],2)
			keydata = split(key,".")
			command = keydata(ubound(keydata))
			Select Case
				Case "currentSource"
					Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
					SetPropertyValue "Multi Room Audio Settings.Zone " & ZStr & " Source", replace(keyvalue[1], chr(34), "")
			End Select	
		End If
	Next
End Sub