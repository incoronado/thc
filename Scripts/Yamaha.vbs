Option Explicit
'On Error Resume Next

Dim Action
Dim SleepVar
'Set su = CreateObject("newObjects.utilctls.StringUtilities")

'-------------------------------------------------------
'- Main: Checks Received Data and Handles Send Data ----
'-------------------------------------------------------
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))

TransmitReadyCommand
SetReportTimeout

Do
	Sleep SleepVar
	Action = GetPropertyValue ("Yamaha V2600 Settings.Action")
	If Action <> "Idle" Then
		MessageHandler(Action)
		SetpropertyValue "Yamaha V2600 Settings.Action", "Idle"
	End If
Loop

Sub MessageHandler(Action)
	Dim a, b
	'Class.Source.Priority.Command:Parameters
    'Example:  Receiver.RS232.10.ProcessSerialMessage:
	a=split(Action,".")
	
    'MsgBox "Ready Command"
    Select Case a(0)
		Case "Receiver"
			' make commands case insensitive; all cases should therefore be expressed as lower case
			b=split(lcase(a(3)),":")
			Select Case b(0)
				Case "placeholder"
				'a
				Case "input"
					'Receiver.RS232.10.Input:Tuner
					SetInput b(1)
				Case "preset"
					'Receiver.RS232.10.Preset:1
					SetPresetStation b(1)
				Case "mastervolumeUp"
          			MasterVolumeUp
				Case "mastervolumedown"
					MasterVolumeDown	
				Case "mastervolumeuprepeat"
					'Receiver.RS232.10.MasterVolumeUpRepeat:5
          			MasterVolumeUpRepeat b(1)
				Case "mastervolumedownrepeat"
					MasterVolumeDownRepeat b(1)		
				End Select
				
			
		' Configuration Command Backwards Compatibility
		Case "MasterPowerToggle"
			MasterPowerToggle
		Case "MasterPowerOn"
			MasterPowerOn
		Case "MasterPowerOff"
			MasterPowerOff
		Case "Zone2PowerOn"
			Zone2PowerOn
		Case "Zone2PowerOff"
			Zone2PowerOff
       Case "Zone2PowerToggle"
	     Zone2PowerToggle
       Case "Zone3PowerToggle"
	     Zone3PowerToggle
	   Case "Zone3PowerOn"
	     Zone3PowerOn
	   Case "Zone3PowerOff"
	     Zone3PowerOff
	   Case "ConvertToLinearZone1"
	     ConvertToLinear(1)  
	   Case "LinearToVolumeZone1"
	     LinearToVolume(1)
	   Case "ConvertToLinearZone2"
	     ConvertToLinear(2)  
	   Case "LinearToVolumeZone2"
	     LinearToVolume(2)
	   Case "ConvertToLinearZone3"
	     ConvertToLinear(3)  
	   Case "LinearToVolumeZone3"
	     LinearToVolume(3)	   
	   Case "Next Input"
	     NextInput
	   	
	   Case "TransmitReadyCommand"
	     TransmitReadyCommand
	   Case "DVD to Computer"
	     TransmitIOAssignCommand
	   Case "SetMasterVolumeToHex"
	     SetVolumeToHex(1)
	   Case "SetZone2VolumeToHex"
	     SetVolumeToHex(2)
 	   Case "SetZone3VolumeToHex"
	     SetVolumeToHex(3)
       Case "MasterVolumeUp"
          MasterVolumeUp
	   Case "MasterVolumeDown"
          MasterVolumeDown
       Case "MasterMuteToggle"
          MasterMuteToggle
	   Case "SetReportTimeout"
	      SetReportTimeout
	   End Select
End Sub


Sub TransmitIOAssignCommand
  Dim SendStr
  SendStr = Chr(20) & "20" & "07" & "0101215" & "0A" 
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 500
End Sub


Sub SetVolumeToHex(zone)
  Dim SendStr
  If zone = 1 Then
     SendStr = Chr(2) & "230" & GetPropertyValue("Yamaha V2600 Settings.AV Master Volume") 
     SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
  ElseIf zone = 2 Then
     SendStr = Chr(2) & "231" & GetPropertyValue("Yamaha V2600 Settings.AV Zone 2 Volume") 
     SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
  ElseIf zone = 3 Then
     SendStr = Chr(2) & "234" & GetPropertyValue("Yamaha V2600 Settings.AV Zone 3 Volume") 
     SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr  
  End If
 'sleep 150
End Sub



Sub MasterPowerOn
  Dim SendStr
  'TransmitReadyCommand
  SendStr = Chr(2) & "07E7E"
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 50
  
 
End Sub

Sub MasterPowerOff
  Dim SendStr
  SendStr = Chr(2) & "07E7F" 
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 50
  
End Sub

Sub MasterPowerToggle
   If GetPropertyValue("Yamaha V2600 Settings.AV Power Master") = "On" Then
      MasterPowerOff
   Else 
      MasterPowerOn
   End If
End Sub


Sub Zone2PowerOn
  Dim SendStr
  SendStr = Chr(2) & "07EBA" 
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 50
End Sub

Sub Zone2PowerOff
  Dim SendStr
  SendStr = Chr(2) & "07EBB" 
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 50
End Sub

Sub Zone2PowerToggle
   If GetPropertyValue("Yamaha V2600 Settings.AV Power Zone 2") = "On" Then
      Zone2PowerOff
   Else 
      Zone2PowerOn
   End If
End Sub


Sub Zone3PowerOn  
  Dim SendStr
  SendStr = Chr(2) & "07AED" 
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 50
End Sub

Sub Zone3PowerOff
  Dim SendStr
  SendStr = Chr(2) & "07AEE" 
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 50
End Sub

Sub Zone3PowerToggle
   If GetPropertyValue("Yamaha V2600 Settings.AV Power Zone 3") = "On" Then
      Zone3PowerOff
   Else 
      Zone3PowerOn
   End If
End Sub


Sub MasterVolumeUpRepeat(RepeatTimes)
	Dim SendStr, i
	for i = 1 to CInt(RepeatTimes)
		SendStr = Chr(2) & "07A1A" 
		SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 		sleep 50
	next		
End Sub

Sub MasterVolumeDownRepeat(RepeatTimes)
	Dim SendStr, i
	for i = 1 to CInt(RepeatTimes)
		SendStr = Chr(2) & "07A1B" 
		SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 		sleep 50
	next		
End Sub


Sub MasterVolumeUp
  Dim SendStr
  SendStr = Chr(2) & "07A1A" 
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 50
End Sub

Sub MasterVolumeDown
  Dim SendStr
  SendStr = Chr(2) & "07A1B" 
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 50
End Sub

Sub MasterMuteToggle
  Dim SendStr
  If GetPropertyValue("Yamaha V2600 Settings.AV Audio Mute") = "On" Then
     SendStr = Chr(2) & "07EA3" 
  Else
     SendStr = Chr(2) & "07EA2" 
  End If
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 150
End Sub

Sub TransmitReadyCommand
	Dim SendStr,  NewReceiveData, i
	SendStr =  Chr(17) & "400"
	' Send a Status Request
	SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr 
	Sleep 500
    'Retry Status Request 4 Times
	For i = 1 to 4
	   NewReceiveData=GetPropertyValue ("Yamaha V2600 Receiver.Received Data")
       
	   If Left(NewReceiveData,1) <> Chr(18) Then
	      'MsgBox Asc(Mid(NewReceiveData,1,1))
		  
		  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
		  Sleep 500
	   End If
	Next
End Sub


Sub SetReportTimeout
  Dim SendStr
  SendStr = Chr(2) & "20102" 
  SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", SendStr
 'sleep 150
End Sub

Sub SetPresetStation (PresetNumber)
	Dim SleepVarPri
	SleepVarPri=100
	
	Select Case PresetNumber
		Case "1"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE0"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE5"
		Case "2"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE0"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE6"
		Case "3"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE0"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE7"
		Case "4"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE0"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE8"
		Case "5"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE0"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE9"
		Case "6"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE0"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AEA"
		Case "7"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE0"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AEB"
		Case "8"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE0"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AEC"
		Case "9"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE1"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE5"
		Case "10"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE1"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE6"
		Case "11"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE1"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE7"
		Case "12"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE1"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE8"
		Case "13"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE1"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE9"
		Case "14"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE1"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AEA"		
		Case "15"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE1"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AEB"
		Case "16"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE1"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AEC"
		Case "17"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE2"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE5"
		Case "18"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE2"
			sleep SleepVarPri
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AE6"
	End Select
	
	Sleep SleepVarPri
End Sub


Sub SetInput(InputName)
	Select Case InputName
		Case "v-aux"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A55" 
		Case "tuner"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A16" 
		Case "cbl/sat"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AC0" 
		Case "dvr/vcr2"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A13" 
		Case "vcr1"	
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A0F"
		Case "dtv"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A54"
		Case "dvd"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AC1"
		Case "md/tape"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A18"
		Case "cd-r"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A19"
		Case "cd"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A15"
		Case "phono"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A14"
		Case "xm"
			SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AB4" 
	End Select


End Sub

Sub NextInput
				If GetPropertyValue("Yamaha V2600 Settings.AV Input") = "V-AUX" Then
				   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A13" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "DVR/VCR2" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A0F" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "VCR1" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AC0" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "CBL/SAT" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A54" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "DTV" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AC1" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "DVD" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A18" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "MD/TAPE" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A19" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "CD-R" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A15" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "CD" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A14" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "PHONO" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A16" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "TUNER" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07AB4" 
				ElseIf GetPropertyValue("Yamaha V2600 Settings.AV Input") = "XM" Then
                   SetPropertyValue "Yamaha V2600 Receiver.Yamaha Serial Command", Chr(02) & "07A55" 
				End If

End Sub


Function HexToDec(strHex)
  dim lngResult
  dim intIndex
  dim strDigit
  dim intDigit
  dim intValue

  lngResult = 0
  for intIndex = len(strHex) to 1 step -1
    strDigit = mid(strHex, intIndex, 1)
    intDigit = instr("0123456789ABCDEF", ucase(strDigit))-1
    if intDigit >= 0 then
      intValue = intDigit * (16 ^ (len(strHex)-intIndex))
      lngResult = lngResult + intValue
    else
      lngResult = 0
      intIndex = 0 ' stop the loop
    end if
  next
  HexToDec = lngResult
End Function



Function VolumeToLinearInt(HexValueStr)
    ' Convert to linear integer value between 0 - 194 
    Dim DecValue
    If HexValueStr = "00" Or HexValueStr = "" Then
       VolumeToLinearInt = 0
	Else
	   VolumeToLinearInt = CInt(HexToDec(HexValueStr)) - 38
    End If
End Function


Sub ConvertToLinear(zone)
   If zone = 1 Then
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Master Volume Dec", VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Master Volume"))
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Master Volume DB", LinearIntToDB(VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Master Volume")))
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Volume 1", Left(LinearIntToDB(VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Master Volume"))),3)
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Volume 2", Right(LinearIntToDB(VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Master Volume"))),1)
     'sleep 5
   ElseIf zone = 2 Then
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Volume Dec", VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Zone 2 Volume"))
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Volume DB", LinearIntToDB(VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Zone 2 Volume")))
     'sleep 5
   ElseIf zone = 3 Then
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Volume Dec", VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Zone 3 Volume"))
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Volume DB", LinearIntToDB(VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Zone 3 Volume")))
     'sleep 5
End If

End Sub

Function LinearIntToDB(LinearInt)
   If LinearInt = 0 Then
      LinearIntToDB = "     "
   Else
      LinearIntToDB = Left(CStr(((LinearInt -1) * .5) - 80) & ".0",5)
   End If
End Function

Function LinearIntToVolume(LinearInt)
    ' Convert to linear integer value between 0 - 194 
    If LinearInt = 0 Then
       LinearIntToVolume = "00"
	Else
	   LinearIntToVolume = Hex(LinearInt + 38)
    End If
End Function

Sub LinearToVolume(zone)
   If zone = 1 Then
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Master Volume", LinearIntToVolume(GetPropertyValue("Yamaha V2600 Settings.AV Master Volume Dec"))
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Master Volume DB", LinearIntToDB(VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Master Volume")))
     'sleep 5
   ElseIf zone = 2 Then
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Volume", LinearIntToVolume(GetPropertyValue("Yamaha V2600 Settings.AV Zone 2 Volume Dec"))
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Volume DB", LinearIntToDB(VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Zone 2 Volume")))
     'sleep 5
   ElseIf zone = 3 Then
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Volume", LinearIntToVolume(GetPropertyValue("Yamaha V2600 Settings.AV Zone 3 Volume Dec"))
     'sleep 5
      SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Volume DB", LinearIntToDB(VolumeToLinearInt(GetPropertyValue("Yamaha V2600 Settings.AV Zone 3 Volume")))
     'sleep 5
   End If
End Sub


