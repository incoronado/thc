Option Explicit
On Error Resume Next

Dim Action
Dim Tuner1ReceivedData, Tuner2ReceivedData
Dim Tuner1OldReceivedData, Tuner2OldReceivedData
Dim SleepVar

'Set su = CreateObject("newObjects.utilctls.StringUtilities")

'-------------------------------------------------------
'- Main: Checks Received Data and Handles Send Data ----
'-------------------------------------------------------
SleepVar = 5

Do
	Sleep 5
	Tuner1ReceivedData = GetPropertyValue ("Tuner 1.Received Data")
	If Tuner1OldReceivedData <> Tuner1ReceivedData Then
		Tuner1OldReceivedData = Tuner1ReceivedData
		Call ReadSerialData(Tuner1ReceivedData,1)
	End If
	
	Tuner2ReceivedData = GetPropertyValue ("Tuner 2.Received Data")
	If Tuner2OldReceivedData <> Tuner2ReceivedData Then
		Tuner2OldReceivedData = Tuner2ReceivedData
		Call ReadSerialData(Tuner2ReceivedData,2)
	End If
	
	
	Action = GetPropertyValue ("Tuner Script.Action")
	If Action <> "Idle" Then
		SetpropertyValue "Tuner Script.Action", "Idle"
		Sleep 5
		If Action <> "" Then
			Call TunerCommand(Action)
		End If
	End If
Loop

Sub ProcessTunerCommand(Action)
	dim a, SendStr
    'SetpropertyValue "UPB Script.Debug", UPBLightID
    'MsgBox "Ready Command"
	' class.command.device
	
	a=split(lcase(Action),".") 
		Select Case a(0)
   			Case "tuner"
				Select Case a(1)

				End Select
		End Select


End Sub

Sub ReadSerialData(Data,tuner)
	Dim SerialMessages, x, i
	 
	If Mid(Data,1,1) = chr(124) Then
		SerialMessages=split(Data, chr(124))	
		for each x in SerialMessages
		'SetpropertyValue "Tuner 1 Settings.Debug", x
			If Mid(x,1,2) = "PS" Then
	           SetpropertyValue "Tuner Script.Debug", "Tuner " & tuner & " Settings.Tuner RDS PS"	
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner RDS PS", Mid(x,4)
			ElseIf Mid(x,1,3) = "PTY" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner RDS PTY", Mid(x,5)
			ElseIf Mid(x,1,2) = "TA" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner RDS TA", Mid(x,4)
			ElseIf Mid(x,1,2) = "TP" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner RDS TP", Mid(x,4)   
			ElseIf Mid(x,1,4) = "FLAG" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner RDS FLAG", Mid(x,1)   
			ElseIf Mid(x,1,4) = "FREQ" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner FREQ", Mid(x,6)   
			ElseIf Mid(x,1,3) = "TID" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner ID", Mid(x,5)
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner RDS RT", ""			   
			ElseIf Mid(x,1,6) = "TUNING" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner TUNING", Mid(x,8)     
			ElseIf Mid(x,1,4) = "MODE" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner MODE", Mid(x,6)
			ElseIf Mid(x,1,5) = "POWER" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner POWER", Mid(x,7)
			ElseIf Mid(x,1,4) = "ZONE" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner ZONE", Mid(x,6)   
			ElseIf Mid(x,1,6) = "PRESET" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner PRESET", Mid(x,7)     
			ElseIf Mid(x,1,2) = "IR" Then  
			   SetpropertyValue "Tuner " & tuner & " Settings.Tuner IR", Mid(x,4)        
			End If
		next		
	Else
		SetpropertyValue "Tuner " & tuner & " Settings.Tuner RDS RT", Mid(Data,1)
			
	End If	


End Sub

