Option Explicit
'On Error Resume Next

Dim Action, SleepVar, ReceivedData, OldReceivedData


SleepVar = 5

Do
	Sleep SleepVar
	ReceivedData = GetPropertyValue ("HDMI Matrix.Received Data")
	If OldReceivedData <> ReceivedData Then
		OldReceivedData = ReceivedData
		Call ReadSerialData(ReceivedData)
	End If
	
	Action = GetPropertyValue ("HDMI Matrix Script.Action")
	If Action <> "Idle" Then
		SetpropertyValue "HDMI Matrix Script.Action", "Idle"
		Sleep SleepVar
		If Action <> "" Then
			Call ProcessMessage(Action)
		End If
	End If
Loop


Sub ProcessMessage(Action)
    'MsgBox "Ready Command"
    Select Case Action 
	' Configuration Command
	   Case "SetA1"
	     SetA1
		Case "SetA2"
	     SetA2
		Case "SetA3"
	     SetA3
		Case "SetA4"
	     SetA4
		Case "SetB1"
	     SetB1
		Case "SetB2"
	     SetB2
		Case "SetB3"
	     SetB3
		Case "SetB4"
	     SetB4 
		 
		 
		 
		 
		Case "PowerToggle" 
		 PowerToggle
		Case "PowerOff"
		 PowerOff
		Case "PowerOn"
		 PowerOn
		 
		 End Select	   
End Sub

Sub ReadSerialData(Data)
	
	If Mid(Data,2,1) = Chr(97) Then
		SetpropertyValue "HDMI Matrix Settings.Debug", "New Command"
		If  Mid(Data,12,1) = Chr(01) Then
			SetpropertyValue "HDMI Matrix Settings.HDMI Power State", "On"
		Else
			SetpropertyValue "HDMI Matrix Settings.HDMI Power State", "Off"
		End if
		
		Select Case Mid(Data,7,1) 
			Case Chr(01)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone A", "1"
			Case Chr(02)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone A", "2"
			Case Chr(04) 
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone A", "3"
			Case Chr(08) 	
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone A", "4"
		End Select
		Sleep SleepVar
		SetpropertyValue "System.Zone A Selected Source", GetPropertyValue("System.Matrix Source " & GetPropertyValue("HDMI Matrix Settings.Matrix Zone A") & " Name")	
		
		Select Case Mid(Data,8,1) 
			Case Chr(01)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone B", "1"
			Case Chr(02)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone B", "2"
			Case Chr(04) 
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone B", "3"
			Case Chr(08) 	
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone B", "4"
		End Select 
		Sleep SleepVar
		SetpropertyValue "System.Zone B Selected Source", GetPropertyValue("System.Matrix Source " & GetPropertyValue("HDMI Matrix Settings.Matrix Zone B") & " Name")	
		
		Select Case Mid(Data,9,1) 
			Case Chr(01)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone C", "1"
			Case Chr(02)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone C", "2"
			Case Chr(04) 
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone C", "3"
			Case Chr(08) 	
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone C", "4"
		End Select 
		Sleep SleepVar
		SetpropertyValue "System.Zone C Selected Source", GetPropertyValue("System.Matrix Source " & GetPropertyValue("HDMI Matrix Settings.Matrix Zone C") & " Name")
		
		Select Case Mid(Data,10,1) 
			Case Chr(01)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone D", "1"
			Case Chr(02)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone D", "2"
			Case Chr(04) 
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone D", "3"
			Case Chr(08) 	
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone D", "4"
		End Select 
		Sleep SleepVar
		SetpropertyValue "System.Zone D Selected Source", GetPropertyValue("System.Matrix Source " & GetPropertyValue("HDMI Matrix Settings.Matrix Zone D") & " Name")	
		
		
		
	End If	
End Sub

Sub PowerOff
	If GetPropertyValue("HDMI Matrix Settings.HDMI Power State") = "On" Then
		PowerToggle
	End If
End Sub

Sub PowerOn
	If GetPropertyValue("HDMI Matrix Settings.HDMI Power State") = "Off" Then
		PowerToggle
	End If
End Sub


Sub PowerToggle
	SetpropertyValue "HDMI Matrix.Serial Command", "10" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "EF" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "D5" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "7B"
End Sub

Sub SetA1
	SetpropertyValue "HDMI Matrix.Serial Command", "00" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "FF" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "D5" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "7B"
End Sub


Sub SetA2
	SetpropertyValue "HDMI Matrix.Serial Command", "01" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "FE" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "D5" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "7B"
End Sub

Sub SetA3
	SetpropertyValue "HDMI Matrix.Serial Command", "02" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "FD" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "D5" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "7B"
End Sub

Sub SetA4
	SetpropertyValue "HDMI Matrix.Serial Command", "03" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "FC" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "D5" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "7B"
End Sub


Sub SetB1
	SetpropertyValue "HDMI Matrix.Serial Command", "04" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "FB" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "D5" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "7B"
End Sub


Sub SetB2
	SetpropertyValue "HDMI Matrix.Serial Command", "05" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "FA" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "D5" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "7B"
End Sub

Sub SetB3
	SetpropertyValue "HDMI Matrix.Serial Command", "06" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "F9" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "D5" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "7B"
End Sub

Sub SetB4
	SetpropertyValue "HDMI Matrix.Serial Command", "07" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "F8" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "D5" 
    sleep 20
	SetpropertyValue "HDMI Matrix.Serial Command", "7B"
End Sub







