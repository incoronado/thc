' Russound
' This script Supports Russound RNET Protocol
' ©2017 Mike Larson created (4/1/2017)
' 
' Requirements:
' RS232 Port COnfigured @ 19200,N,8,1 No Flow Control
'
' Limitations:
' 
' 
Option Explicit
'On Error Resume Next

Dim Action
Dim SleepVar

SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))

Do
	Action = GetPropertyValue ("Multiroom Audio Settings.Action")
	If Action <> "Idle" Then
		MessageHandler(Action)
		SetpropertyValue "Multiroom Audio Settings.Action", "Idle"
	End If
	Sleep SleepVar
Loop

Sub MessageHandler(Action)
	Dim a, b
	'Class.Source.Priority.Command:Parameters
    'Example:  Receiver.RS232.10.ProcessSerialMessage:
	a=split(Action,".")
	
    'MsgBox "Ready Command"
    Select Case a(0)
		Case "Russound"

	End Select
	
End Sub		