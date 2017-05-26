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
			ResponseType = Left(keyvalue(0),1)
			key = Mid(keyvalue(0),3)
			If Instr(key,".") Then
				keydata = split(key,".")
				command = keydata(ubound(keydata))
			Else 
				command = ""	
			End if	
			
			SetPropertyValue "Multiroom Audio Settings.Debug", command

			Select Case command
				Case "currentSource"
					Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
					SetPropertyValue "Multiroom Audio Settings.Zone " & ZStr & " Source", replace(keyvalue(1), chr(34), "")
					SetPropertyValue "Multiroom Audio Settings.Debug 2", ZStr
				Case "mode"	
					Zstr = replace(replace(replace(keydata(0),"[",""),"]",""),"S","")
					'SetPropertyValue "Multiroom Audio Settings.Debug 2", replace(keyvalue(1), chr(34),"")
					If CInt(ZStr) = 2 Then
						'SetPropertyValue "Russound.Streamer Source", replace(keyvalue(1), chr(34),"")
					End If		
			End Select	
		End If
	Next
End Sub