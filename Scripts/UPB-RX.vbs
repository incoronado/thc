Option Explicit
On Error Resume Next

Dim DataItem, SleepVar
DataItem = GetPropertyValue ("UPB PIM.Received Data")
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
Select Case Mid(DataItem, 1, 2)
	' Configuration Command DC2
	Case "PU"
		Select Case Mid(DataItem, 1, 4) & Mid(DataItem, 13, 2)
			Case "PU0886"
				If Mid(DataItem, 15, 2) = "00" Then
					SetpropertyValue "UPB-" + Mid(DataItem, 11, 2) + ".Power State", "Off"
				Else
					SetpropertyValue "UPB-" + Mid(DataItem, 11, 2) + ".Power State", "On"
				End if
				SetpropertyValue "UPB-" + Mid(DataItem, 11, 2) + ".Dim Level", CLng("&h" & Mid(DataItem, 15, 2))
				UpdateRemoteData DataItem
			Case "PU8922"
				'SetpropertyValue "UPB Script.Action", "upblight.poll"
		End Select		

	Case "PA"
			
	Case "PK"
		
	Case "PB"
		
	Case "PN"
		
	Case "PE"
	
End Select

Sleep SleepVar

Sub UpdateRemoteData(DataItem)
	Dim RemoteList, i
	RemoteList = split(GetPropertyValue("System.HBRemoteList"),",")
	
	For Each i in RemoteList
	    
		If Mid(DataItem, 15, 2) = "00" Then	
						
			If GetPropertyValue("Remote-" & CStr(i) & ".UPB Light ID") = Mid(DataItem, 11, 2) Then
			    'SetPropertyValue "System.Debug", "Got Here"
				SetPropertyValue "Remote-" & CStr(i) & ".UPB Light State", "Off"
			End if
		Else
			If GetPropertyValue("Remote-" & CStr(i) & ".UPB Light ID") = Mid(DataItem, 11, 2) Then
				SetPropertyValue "Remote-" & CStr(i) & ".UPB Light State", "On"
			End if
		End if
		If GetPropertyValue("Remote-" & CStr(i) & ".UPB Light ID") = Mid(DataItem, 11, 2) Then
			SetPropertyValue "Remote-" & CStr(i) & ".UPB Dim Level", CLng("&h" & Mid(DataItem, 15, 2))
		End if
	Next
End Sub

