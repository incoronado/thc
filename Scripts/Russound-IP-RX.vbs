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
					UpdateRemoteData
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
					UpdateRemoteData
				Case "mode"
					If Mid(keydata(0), 1, 2) = "S[" Then
					    SetPropertyValue "Russound.mode", replace(keyvalue(1),chr(34),"")
					End if	
				Case "playlistName"
					If Mid(keydata(0), 1, 2) = "S[" Then
						SetPropertyValue "Russound.playlistName", replace(keyvalue(1),chr(34),"")
					End if	
				Case "channelName"
					If Mid(keydata(0), 1, 2) = "S[" Then
						SetPropertyValue "Russound.channelName", replace(keyvalue(1),chr(34),"")
					End if	
				Case "albumName"
					If Mid(keydata(0), 1, 2) = "S[" Then
						SetPropertyValue "Russound.albumName", replace(keyvalue(1),chr(34),"")
					End if	
				Case "artistName"
					If Mid(keydata(0), 1, 2) = "S[" Then
						SetPropertyValue "Russound.artistName", replace(keyvalue(1),chr(34),"")
					End if	
				Case "songName"
					If Mid(keydata(0), 1, 2) = "S[" Then
						SetPropertyValue "Russound.songName", replace(keyvalue(1),chr(34),"")
					End if			
				Case "coverArtURL"
					If Mid(keydata(0), 1, 2) = "S[" Then
						Zstr = replace(replace(replace(keydata(1),"[",""),"]",""),"Z","")
						SetPropertyValue "Russound.coverArtURL", replace(keyvalue(1),chr(34),"")
					End if			
				Case "repeatMode"
					If Mid(keydata(0), 1, 2) = "S[" Then
						SetPropertyValue "Russound.repeatMode", treplace(keyvalue(1),chr(34),"")
					End if			
				Case "shuffleMode"
					If Mid(keydata(0), 1, 2) = "S[" Then
						SetPropertyValue "Russound.shuffleMode", replace(keyvalue(1),chr(34),"")
					End if
				Case "channel"
					If Mid(keydata(0), 1, 2) = "S[" Then
						SetPropertyValue "Russound.channel", replace(keyvalue(1),chr(34),"")
					End if	
				Case "programServiceName"
					If Mid(keydata(0), 1, 2) = "S[" Then
						SetPropertyValue "Russound.programServiceName", replace(keyvalue(1),chr(34),"")
					End if	
				Case "radioText"
					If Mid(keydata(0), 1, 2) = "S[" Then
						SetPropertyValue "Russound.radioText", replace(keyvalue(1),chr(34),"")
					End if	

			End Select	
		End If
	Next
End Sub


Sub UpdateRemoteData
	Dim i, SelectedZone, x, ThemesFolder, fso

	Set fso = CreateObject("Scripting.FileSystemObject")


	For i = 1 to 4
	    ThemesFolder =  GetPropertyValue("Remote-" & CStr(i) &  ".Themes Folder")
		SelectedZone = GetPopropertyValue("Remote-" & CStr(i) & ".Selected Zone")
		For x = 1 to 8
	   		If GetPropertyValue("Multiroom Audio Settings.Zone " & Cstr(x) " Power")  = "On" Then
	   			If SelectZone = CStr(x) Then
	   			    ' Check to see if file exists. There might not be one
	   				If fso.FileExists("Config\Themes\" & ThemesFolder & "\icons\zone" & CStr(x) & "-sel-on.png") Then
	   					SetPropertyValue "Remote-" & CStr(i) & ".Menu Icon " & Cstr(x), "Config\Themes\" & ThemesFolder & "\icons\zone" & CStr(x) & "-sel-on.png"
	   				End If	
	   			Else
	   				If fso.FileExists("Config\Themes\" & ThemesFolder & "\icons\zone" & CStr(x) & "-unsel-on.png") Then
	   					SetPropertyValue "Remote-" & CStr(i) & ".Menu Icon " & Cstr(x), "Config\Themes\" & ThemesFolder & "\icons\zone" & CStr(x) & "-unsel-on.png"
	   				End If	
	   			End if
	   		Else
	   			If SelectZone = CStr(x) Then
	   				If fso.FileExists("Config\Themes\" & ThemesFolder & "\icons\zone" & CStr(x) & "-sel-off.png") Then
	   					SetPropertyValue "Remote-" & CStr(i) & ".Menu Icon " & Cstr(x), "Config\Themes\" & ThemesFolder & "\icons\zone" & CStr(x) & "-sel-off.png"
	   				End If	
	   			Else
	   				If fso.FileExists("Config\Themes\" & ThemesFolder & "\icons\zone" & CStr(x) & "-unsel-off.png") Then 
	   					SetPropertyValue "Remote-" & CStr(i) & ".Menu Icon " & Cstr(x), "Config\Themes\" & ThemesFolder & "\icons\zone" & CStr(x) & "-unsel-off.png"
	   				End If	
	   			End if
	   		End if
	   	Next


  	Next
 
  	Set fso = Nothing

End Sub
