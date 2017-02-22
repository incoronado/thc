Option Explicit
Dim Action, Sleepvar, Status, CursorRatio

Sleepvar = 5

Do
	Action = GetPropertyValue ("Keyboard Script.Action")
	If Action <> "Idle" Then
		SetpropertyValue "Keyboard Script.Action", "Idle"
		If Action <> "" Then
			Call Message_Handler(Action)
			If Err.Number <> 0 Then
				Status = "Error: " & Err.Description
				Err.Clear
			End If
		End If
	End If
	Sleep Sleepvar
Loop

Sub Message_Handler(Action)
	Dim a, CursorStr, KbdStr 
	a=split(Action,".")
	CursorStr = ""
	KbdStr=GetPropertyValue("GalaxyProTab1 System." & a(1))
	'Keyboard.Jukebox - Search.Sendkey.Q
	'Keyboard.SendKey.a.Jukebox
	Select Case a(0)
		Case "Keyboard"					
			Select Case a(2)
				Case "Sendkey"
					If GetPropertyValue ("Keyboard.Shift State") = "Off" Then
						KbdStr = KbdStr & LCase(a(3))
					Else
						KbdStr = KbdStr & UCase(a(3))
					End if
				Case "Esc"
					KbdStr = ""
					CloseRemotePanel "Keyboard"
				Case "Enter"
					If a(1) = "Jukebox - Search" Then
						SetPropertyValue "Jukebox.Action", "Song Select No Drilldown"
						CloseRemotePanel "Keyboard"
					ElseIf a(1) = "Jukebox - New Playlist" Then	
						SetPropertyValue "Jukebox.Action", "Create Playlist"
						CloseRemotePanel "Keyboard"
					End If
				Case "Shift"
					If GetPropertyValue ("Keyboard.Shift State") = "Off" Then
						SetPropertyValue "Keyboard.Shift State", "On"
					Else
						SetPropertyValue "Keyboard.Shift State", "Off"
					End if
				Case "Backspace"
					If Len(KbdStr) < 2 Then
						KbdStr = "" 
					Else
						KbdStr = Left(KbdStr, Len(KbdStr) -1)
					End if 
				Case "Spacebar"
					KbdStr = KbdStr & " "
				Case "Semicolon"
					KbdStr = KbdStr & chr(59)	
				Case "Period"
					KbdStr = KbdStr & chr(46)	
				Case "Comma"
					KbdStr = KbdStr & chr(44)
			End Select
			
	End Select 
	
	 CursorRatio = CInt((Len(KbdStr)) * 4.20)
	  ' MsgBox CursorRatio
	  CursorStr = Space(CursorRatio)
	  CursorStr = CursorStr & "^"
	  SetPropertyValue "Keyboard.Keyboard Cursor", CursorStr
	  SetPropertyValue "GalaxyProTab1 System." & a(1), KbdStr
	  
End Sub



