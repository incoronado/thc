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
					If Mid(keydata(0), 1, 2) = "C[" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						If replace(keyvalue(1), chr(34), "") = "ON" Then
							SetPropertyValue "Multiroom Audio Settings.Zone " & trim(ZStr) & " Power", "On"
						Else
							SetPropertyValue "Multiroom Audio Settings.Zone " & trim(ZStr) & " Power", "Off"
						End If	
					ElseIf Mid(keydata(0), 1, 2) = "S[" Then
						If replace(keyvalue(1), chr(34), "") = "ON" Then
							SetPropertyValue "Multiroom Audio Settings.Source " & trim(ZStr) & " Power", "On"
						Else
							SetPropertyValue "Multiroom Audio Settings.Source " & trim(ZStr) & " Power", "Off"	
						End if	
					End if
				Case "mode"
					If Mid(keydata(0), 1, 2) = "S[2" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Russound.mode", trim(ZStr)
					End if	
				Case "playlistName"
					If Mid(keydata(0), 1, 2) = "S[2" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Russound.playlistName", trim(ZStr)
					End if	
				Case "channelName"
					If Mid(keydata(0), 1, 2) = "S[2" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Russound.channelName", trim(ZStr)
					End if	
				Case "albumName"
					If Mid(keydata(0), 1, 2) = "S[2" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Russound.albumName", trim(ZStr)
					End if	
				Case "artistName"
					If Mid(keydata(0), 1, 2) = "S[2" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Russound.artistName", trim(ZStr)
					End if	
				Case "songName"
					If Mid(keydata(0), 1, 2) = "S[2" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Russound.songName", trim(ZStr)
					End if			
				Case "coverArtURL"
					If Mid(keydata(0), 1, 2) = "S[2" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Russound.coverArtURL", trim(ZStr)
					End if			
				Case "repeatMode"
					If Mid(keydata(0), 1, 2) = "S[2" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Russound.repeatMode", trim(ZStr)
					End if			
				Case "shuffleMode"
					If Mid(keydata(0), 1, 2) = "S[2" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Russound.shuffleMode", trim(ZStr)
					End if					
			End Select	
		End If
	Next
End Sub