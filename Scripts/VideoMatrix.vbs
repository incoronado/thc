Option Explicit
'On Error Resume Next

Dim Action, SleepVar

SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
Do
	Sleep SleepVar	
	Action = GetPropertyValue ("HDMI Matrix Script.Action")
	If Action <> "Idle" Then
		ProcessMessage(Action)
		SetpropertyValue "HDMI Matrix Script.Action", "Idle"
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
