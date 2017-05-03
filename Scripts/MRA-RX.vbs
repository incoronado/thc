Option Explicit
'On Error Resume Next

Dim Data, SleepVar
Data = GetPropertyValue ("Multiroom Audio Amplifier.Received Data")
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))


'SetPropertyValue "Multiroom Audio Settings.Debug", "Got MRA"
Dim CMDStr
Select Case Mid(Data, 1, 3)
	Case chr(10) & "#<"
		Select Case Mid(Data, 6, 2)
			Case "PR"
			Case "MU"
			Case "DT"
			Case "TR"
			Case "VO"
				UpdateOnChange "Volume", Mid(Data, 5, 1), CStr(CInt(Mid(Data, 8, 2)))
			Case "TR"
			Case "BS"
			Case "BL"
			Case "CH"
			Case "LS"
		End Select
	Case chr(10) & "#>"
		Select Case Mid(Data, 4, 2)
			Case "PR"
			Case "MU"
			Case "DT"
			Case "TR"
			Case "VO"
			Case "TR"
			Case "BS"
			Case "BL"
			Case "CH"
			Case "LS"
			
			Case Else
				If Mid(Data, 6, 2)  = "00" Then
					CMDStr = "Off"
				ElseIf Mid(Data, 6, 2)  = "01" Then	
					CMDStr = "On"
				End if
				UpdateOnChange "PA", Mid(Data, 5, 1), CMDStr
			
				If Mid(Data, 8, 2)  = "00" Then
					CMDStr = "Off"
				ElseIf Mid(Data, 8, 2)  = "01" Then	
					CMDStr = "On"
				End if	
				UpdateOnChange "Power", Mid(Data, 5, 1), CMDStr
				
				If Mid(Data, 10, 2)  = "00" Then
					CMDStr = "Off"
				ElseIf Mid(Data, 10, 2)  = "01" Then	
					CMDStr = "On"
				End if
				UpdateOnChange "Mute", Mid(Data, 5, 1) , CMDStr
				UpdateOnChange "Volume", Mid(Data, 5, 1), CStr(CInt(Mid(Data, 14, 2)))
				UpdateOnChange "Treble", Mid(Data, 5, 1), CStr(CInt(Mid(Data, 16, 2)))
				UpdateOnChange "Bass", Mid(Data, 5, 1), CStr(CInt(Mid(Data, 18, 2)))
				UpdateOnChange "Balance", Mid(Data, 5, 1), CStr(CInt(Mid(Data, 20, 2)))
				UpdateOnChange "Source" , Mid(Data, 5, 1), CStr(CInt(Mid(Data, 22, 2)))
		End Select
End Select


Sub UpdateOnChange (PropertyType, PropertyZone, PropertyValue)
	Dim PropertyName, Remotelist, i
	RemoteList = split(GetPropertyValue("System.HBRemoteList"),",")
	PropertyName = "Multiroom Audio Settings.Zone " & CStr(PropertyZone) & " " & PropertyType
		SetpropertyValue PropertyName, PropertyValue
		For Each i in RemoteList
			If  GetPropertyValue("Remote-" & CStr(i) & ".Selected Zone") = CStr(PropertyZone) Then
				Select Case PropertyType
					Case "Power"
						SetPropertyValue "Remote-" & CStr(i) & ".Selected Zone Power", PropertyValue
					Case "Source"
						SetPropertyValue "Remote-" & CStr(i) & ".Selected Source", PropertyValue
					Case "Volume"
						SetPropertyValue "Remote-" & CStr(i) & ".Selected Zone Volume", PropertyValue
				End Select	
			End if
		Next

End Sub
