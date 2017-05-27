Option Explicit
'On Error Resume Next


Dim Data, SleepVar
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
ReadIPData GetPropertyValue("Russound IP.Received Data")
'Sleep SleepVar


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
			
			If command <> "" Then
				SetPropertyValue "Multiroom Audio Settings.Debug", command
			End If

			Select Case command
				Case "currentSource"
					Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
					SetPropertyValue "Multiroom Audio Settings.Zone " & ZStr & " Source", replace(keyvalue(1), chr(34), "")
				Case "mode"	
					Zstr = replace(replace(replace(keydata(0),"[",""),"]",""),"S","")
					If CInt(ZStr) = 2 Then
						SetPropertyValue "Russound.Streamer Source", replace(keyvalue(1), chr(34),"")
					End If
				Case "volume"
				    If Mid(keydata(0), 1, 2) = "C[" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Multiroom Audio Settings.Zone " & trim(ZStr) & " Volume", replace(keyvalue(1), chr(34), "")
					ElseIf Mid(keydata(0), 1, 2) = "S[" Then
							SetPropertyValue "Multiroom Audio Settings.Source " & trim(ZStr) & " Volume", replace(keyvalue(1), chr(34), "")

					End if
				Case "status"
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Multiroom Audio Settings.Zone " & trim(ZStr) & " Power", replace(keyvalue(1), chr(34), "")

			End Select	
		End If
	Next
End Sub