Option Explicit
'On Error Resume Next

Dim Data, SleepVar
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
ReadSerialData GetPropertyValue("Yamaha V2600 Receiver.Received Data")

Sleep SleepVar

Sub ReadSerialData(Data)
	Dim RptStatusNbr, RptCmdTyp, RptCmdNbr, RptItem, DataItem, DataBuffer, msglength
	'SetPropertyValue "Yamaha V2600 Settings.AV Debug", data
	DataItem = Data
   	 		Select Case Left(DataItem, 1)
	
		    ' Configuration Command DC2
			Case Chr(18)
				If Mid(DataItem,8,2) = "09" Then
					RptCmdNbr=Mid(DataItem,18,1)
					If RptCmdNbr = "0" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					ElseIf RptCmdNbr = "1" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					ElseIf RptCmdNbr = "2" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					ElseIf RptCmdNbr = "3" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					ElseIf RptCmdNbr = "4" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					ElseIf RptCmdNbr = "5" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					ElseIf RptCmdNbr = "6" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					ElseIf RptCmdNbr = "7" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					End If					
			    End If
   			   ' Ignore if a Ready Command
			   If Mid(DataItem,8,2) = "91" Then
				   SetPropertyValue "Yamaha V2600 Settings.AV Model ID", Mid(DataItem,2,5)
				   'sleep SleepVar
				   SetPropertyValue "Yamaha V2600 Settings.AV Version", Mid(DataItem,7,1)
				   'sleep SleepVar
				   msglength = HexToDec(Mid(DataItem,8,2))
				   ' System
				   RptCmdNbr=Mid(DataItem,17,1)
				   If RptCmdNbr = "0" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV System", "OK"
					ElseIf RptCmdNbr = "1" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV System", "Busy"
				   ElseIf RptCmdNbr = "2" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV System", "Standby"
				   End If
				   'sleep SleepVar
				   RptCmdNbr=Mid(DataItem,18,1)
				   'Power
				   'msgbox RptCmdNbr
				   If RptCmdNbr = "0" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
				   ElseIf RptCmdNbr = "1" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
				   ElseIf RptCmdNbr = "2" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					ElseIf RptCmdNbr = "3" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					ElseIf RptCmdNbr = "4" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					ElseIf RptCmdNbr = "5" Then
  
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					ElseIf RptCmdNbr = "6" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					ElseIf RptCmdNbr = "7" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
					   SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					End If
					'sleep SleepVar

					' Zone 1 Input
					RptCmdNbr=Mid(DataItem,19,1)
					If RptCmdNbr = "0" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "PHONO"
					ElseIf RptCmdNbr = "1" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "CD"
					ElseIf RptCmdNbr = "2" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "TUNER"
					ElseIf RptCmdNbr = "3" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "CD-R"
					ElseIf RptCmdNbr = "4" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "MD/TAPE"	  
					ElseIf RptCmdNbr = "5" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "DVD"	  
					ElseIf RptCmdNbr = "6" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "DTV"	  
					ElseIf RptCmdNbr = "7" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "CBL/SAT"	  
					ElseIf RptCmdNbr = "8" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "SAT"
					ElseIf RptCmdNbr = "9" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "VCR1"			
					ElseIf RptCmdNbr = "A" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "DVR/VCR2"			
					ElseIf RptCmdNbr = "B" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "VCR3/DVR"			
					ElseIf RptCmdNbr = "C" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "V-AUX"			
					ElseIf RptCmdNbr = "E" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Input", "XM"			
					End If
					'sleep SleepVar
					
					' Multi Channel Input
					RptCmdNbr=Mid(DataItem,20,1)
					
					'sleep SleepVar
					'Audio Select
					RptCmdNbr=Mid(DataItem,21,1)
					If RptCmdNbr = "0" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Audio Select", "Auto"
					ElseIf RptCmdNbr = "3" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Audio Select", "COAX/OPT"
					ElseIf RptCmdNbr = "4" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Audio Select", "Analog"
					ElseIf RptCmdNbr = "5" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Audio Select", "Analog Only"
					ElseIf RptCmdNbr = "8" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Audio Select", "HDMI"	 
					End if
					'sleep SleepVar

					'Audio Mute
					RptCmdNbr=Mid(DataItem,22,1)
					If RptCmdNbr = "0" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Audio Mute", "Off"
					ElseIf RptCmdNbr = "3" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Audio Mute", "On"
					End If
					
					'sleep SleepVar

					' Zone 2 Input
					RptCmdNbr=Mid(DataItem,23,1)
					If RptCmdNbr = "0" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "PHONO"
					ElseIf RptCmdNbr = "1" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "CD"
					ElseIf RptCmdNbr = "2" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "TUNER"
					ElseIf RptCmdNbr = "3" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "CD-R"
					ElseIf RptCmdNbr = "4" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "MD/TAPE"	  
					ElseIf RptCmdNbr = "5" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "DVD"	  
					ElseIf RptCmdNbr = "6" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "DTV"	  
					ElseIf RptCmdNbr = "7" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "CBL/SAT"	  
					ElseIf RptCmdNbr = "8" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "SAT"
					ElseIf RptCmdNbr = "9" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "VCR1"			
					ElseIf RptCmdNbr = "A" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "DVR/VCR2"			
					ElseIf RptCmdNbr = "C" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "V-AUX"			
					ElseIf RptCmdNbr = "E" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "XM"			
					End If
					'sleep SleepVar

				   ' Zone 2 Mute
					RptCmdNbr=Mid(DataItem,24,1)
					If RptCmdNbr = "0" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Mute", "Off"
					ElseIf RptCmdNbr = "1" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Mute", "On"
					End If
					'sleep SleepVar

				   ' Master Volume 2 Byte Value
					SetPropertyValue "Yamaha V2600 Settings.AV Master Volume", Mid(DataItem,25,2)
					'sleep SleepVar
					ConvertToLinear(1)
					

				   'Zone 2 Volume 2 Byte Value
				   SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Volume", Mid(DataItem,27,2)
				   'sleep SleepVar
				   ConvertToLinear(2)

				   'Zone 3 Volume 3 Byte Value
				   SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Volume", Mid(DataItem,139,2)
				   'sleep SleepVar
				   ConvertToLinear(3)
				   
				   ' Program 2 Byte Value
				   SetPropertyValue "Yamaha V2600 Settings.AV Program 2 Byte", Mid(DataItem,29,2)
				   'sleep SleepVar
				   
				   
				   

				   'Effect
					RptCmdNbr=Mid(DataItem,31,1)
					
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Effect", "Off"
					Else
						SetPropertyValue "Yamaha V2600 Settings.AV Effect", "On"
					End If
					'sleep SleepVar

				   'Extended Surround
					'MsgBox Mid(DataItem,32,1)
					RptCmdNbr=Mid(DataItem,32,1)
					If RptCmdNbr = "0" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "Off"
					ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "EX/ES"
					ElseIf RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "Discrete On"
					ElseIf RptCmdNbr = "3" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "Auto"
					ElseIf RptCmdNbr = "4" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "EX"	  
					ElseIf RptCmdNbr = "5" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "PLIIx Movie"	  
					ElseIf RptCmdNbr = "6" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "PLIIx Music"
					End If

					'sleep SleepVar

				   ' OSD
					RptCmdNbr=Mid(DataItem,33,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV OSD", "Full"	
					Elseif RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV OSD", "Short"	
					Elseif RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV OSD", "Off"	
					End If
					'sleep SleepVar
					
					' Sleep
					RptCmdNbr=Mid(DataItem,34,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Sleep", "120"	
					Elseif RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Sleep", "90"	
					Elseif RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Sleep", "60"
					Elseif RptCmdNbr = "3" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Sleep", "30"
					Elseif RptCmdNbr = "4" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Sleep", "Off"
					End If
					'sleep SleepVar
					' Sleep
					RptCmdNbr=Mid(DataItem,35,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Page", "A"	
					Elseif RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Page", "B"	
					Elseif RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Page", "C"
					Elseif RptCmdNbr = "3" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Page", "D"
					Elseif RptCmdNbr = "4" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Page", "E"
					End If
					'sleep SleepVar
					
					RptCmdNbr=Mid(DataItem,36,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "1"	
					Elseif RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "2"	
					Elseif RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "3"
					Elseif RptCmdNbr = "3" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "4"
					Elseif RptCmdNbr = "4" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "5"
					Elseif RptCmdNbr = "5" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "6"
					Elseif RptCmdNbr = "6" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "7"
					Elseif RptCmdNbr = "7" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "8"
					End If
					'sleep SleepVar
					
					'Night Mode
					RptCmdNbr=Mid(DataItem,37,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Night Mode", "Off"	
					Elseif RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Night Mode", "Cinema"	
					Elseif RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Night Mode", "Music"
					End If
					'Sleep SleepVar
					
					'Night Mode Parameter
					RptCmdNbr=Mid(DataItem,38,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Night Mode Parameter", "Low"	
					Elseif RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Night Mode Parameter", "Mid"	
					Elseif RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Night Mode Parameter", "High"
					End If
					'Sleep SleepVar
					
					'Speaker A
					RptCmdNbr=Mid(DataItem,39,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Speaker A", "Off"	
					Elseif RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Speaker A", "On"	
					End If
					'Sleep SleepVar
					
					'Speaker B
					RptCmdNbr=Mid(DataItem,40,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Speaker B", "Off"	
					Elseif RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Speaker B", "On"	
					End If
					'Sleep SleepVar
					
					'Playback
					RptCmdNbr=Mid(DataItem,41,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "Multi CH Input"
					 ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "Analog"
					 ElseIf RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "PCM"
					 ElseIf RptCmdNbr = "3" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "D.D.(except for 2/0)"
					 ElseIf RptCmdNbr = "4" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "D.D.(2/0)"
					 ElseIf RptCmdNbr = "5" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "D.D. Karaoke"
					 ElseIf RptCmdNbr = "6" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "D.D.EX"
					 ElseIf RptCmdNbr = "7" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "DTS"
					 ElseIf RptCmdNbr = "8" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "DTS. ES"
					 ElseIf RptCmdNbr = "9" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "Other Digital"
					 ElseIf RptCmdNbr = "A" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "DTS Analog Mute"
					 ElseIf RptCmdNbr = "B" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "DTS Discrete"
					 ElseIf RptCmdNbr = "C" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "Other than AAC 2/0"
					 ElseIf RptCmdNbr = "D" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "AAC 2/0"
					 End if
					 'Sleep SleepVar
					 
					 ' FS
					 RptCmdNbr=Mid(DataItem,42,1)
					 If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "Analog"
					 ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "32kHz"
					 ElseIf RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "44.1kHz"
					 ElseIf RptCmdNbr = "3" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "48kHz"
					 ElseIf RptCmdNbr = "4" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "64kHz"
					 ElseIf RptCmdNbr = "5" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "88.2kHz"
					 ElseIf RptCmdNbr = "6" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "96kHz"
					 ElseIf RptCmdNbr = "7" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "Unknown"
					 ElseIf RptCmdNbr = "8" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "128kHz"
					 ElseIf RptCmdNbr = "9" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "176.4.kHz"
					 ElseIf RptCmdNbr = "A" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "192.0khz"
					 ElseIf RptCmdNbr = "B" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "DTS 96/24"
					 End if
					'Sleep SleepVar
					
					'AV EX/ES
					 RptCmdNbr=Mid(DataItem,43,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EX/ES", "Off"
					 ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EX/ES", "Matrix On"
					 ElseIf RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EX/ES", "Discrete On"
					 End If
					'Sleep SleepVar
					
					'THR / Bypass
					RptCmdNbr=Mid(DataItem,44,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV THR/Bypass", "Off"
					 ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV THR/Bypass", "On"
					 End If
					 'Sleep SleepVar
					 
					'RED DTS
					RptCmdNbr=Mid(DataItem,45,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Red DTS", "Release"
					 ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Red DTS", "Wait"
					 End If
					 'Sleep SleepVar
					
					'Head Phone
					RptCmdNbr=Mid(DataItem,46,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Head Phone", "Off"
					 ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Head Phone", "On"
					 End If
					 'Sleep SleepVar
					 
					 'Tuner Band
					RptCmdNbr=Mid(DataItem,47,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Band", "FM"
					 ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Band", "AM"
					 End If
					 'Sleep SleepVar
					 
					'Tuner Tuned
					RptCmdNbr=Mid(DataItem,48,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Tuned", "NOT Tuned"
					 ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Tuned", "Tuned"
					 End If
					 'Sleep SleepVar
					 
					'DC1 Trigger Output
					RptCmdNbr=Mid(DataItem,49,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV DC1 Trigger Output", "Off"
					 ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV DC1 Trigger Output", "On"
					 End If
					 'Sleep SleepVar
					 
					'AV Decoder Mode
					 RptCmdNbr=Mid(DataItem,50,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Decoder Mode", "Auto"
					 ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Decoder Mode", "DTS"
					 End If
					 'Sleep SleepVar
					 
					 'AV Dual Mono
					 RptCmdNbr=Mid(DataItem,51,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Dual Mono", "Main"
					ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Dual Mono", "Sub"
					ElseIf RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Dual Mono", "All"
					ElseIf RptCmdNbr = "8" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Dual Mono", "N/A"
					 End If
					 'Sleep SleepVar
					 
					  'DC1 Trigger Control
					 RptCmdNbr=Mid(DataItem,52,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV DC1 Trigger Control", "All Zone"
					ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV DC1 Trigger Control", "Main"
					ElseIf RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV DC1 Trigger Control", "Zone 2"
					ElseIf RptCmdNbr = "3" Then
						SetPropertyValue "Yamaha V2600 Settings.AV DC1 Trigger Control", "Zone 3"
					 End If
					 'Sleep SleepVar
					 
					'DTS 96/24
					RptCmdNbr=Mid(DataItem,53,1)
					If RptCmdNbr = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV DTS 96/24", "Off"
					ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV DTS 96/24", "On"
					End If
					'Sleep SleepVar
					
					
					 
				    'Zone 2 Input
					RptCmdNbr=Mid(DataItem,129,1)
					If RptCmdNbr = "0" Then
					   SetPropertyValue "Yamaha V2600 Settings.AV Decoder Select", "Pro Logic"
					ElseIf RptCmdNbr = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Decoder Select", "PLIIx Movie"
					ElseIf RptCmdNbr = "2" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Decoder Select", "PLIIx Music"
					ElseIf RptCmdNbr = "3" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Decoder Select", "PLIIx Game"
					ElseIf RptCmdNbr = "4" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Decoder Select", "Neo:6 Cinema"	  
					ElseIf RptCmdNbr = "5" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Decoder Select", "Neo:6 Music"	  
					End If
					'sleep SleepVar
				End if

            ' System Command DC1
			Case Chr(17)
			   RptCmdNbr=Mid(DataItem,2,2)
			   Select Case RptCmdNbr
                  Case "00"
				     ' Tuner Frequency
                     SetPropertyValue "Yamaha V2600 Settings.AV Tuner Frequency", Mid(DataItem,4,7)
               End Select         


  		   ' System Status Report STX
			Case Chr(02)

			   RptCmdTyp=Mid(DataItem,2,1)
			   RptCmdNbr=Mid(DataItem,4,2)
			   RptStatusNbr=Mid(DataItem,6,2)
		       
			   Select Case RptCmdNbr
				  Case "10"
					 RptItem = "Playback"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "Multi CH Input"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "Analog"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "PCM"
					 ElseIf RptStatusNbr = "03" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "D.D.(except for 2/0)"
					 ElseIf RptStatusNbr = "04" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "D.D.(2/0)"
					 ElseIf RptStatusNbr = "05" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "D.D. Karaoke"
					 ElseIf RptStatusNbr = "06" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "D.D.EX"
					 ElseIf RptStatusNbr = "07" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "DTS"
					 ElseIf RptStatusNbr = "08" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "DTS. ES"
					 ElseIf RptStatusNbr = "09" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "Other Digital"
					 ElseIf RptStatusNbr = "0A" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "DTS Analog Mute"
					 ElseIf RptStatusNbr = "0B" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "DTS Discrete"
					 ElseIf RptStatusNbr = "0C" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "Other than AAC 2/0"
					 ElseIf RptStatusNbr = "0D" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Playback Decoding", "AAC 2/0"
					 End if
				  Case "11"
					 RptItem = "FS"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "Analog"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "32kHz"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "44.1kHz"
					 ElseIf RptStatusNbr = "03" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "48kHz"
					 ElseIf RptStatusNbr = "04" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "64kHz"
					 ElseIf RptStatusNbr = "05" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "88.2kHz"
					 ElseIf RptStatusNbr = "06" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "96kHz"
					 ElseIf RptStatusNbr = "07" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "Unknown"
					 ElseIf RptStatusNbr = "08" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "128.0kHz"
					 ElseIf RptStatusNbr = "09" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "176.4.kHz"
					 ElseIf RptStatusNbr = "0A" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "192.0khz"
					 ElseIf RptStatusNbr = "0B" Then
						SetPropertyValue "Yamaha V2600 Settings.AV FS", "DTS 96/24"
					 End if
				
				  Case "12"
					 RptItem = "ES/EX"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EX/ES", "Off"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EX/ES", "Matrix On"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EX/ES", "Discrete On"
					 End If
				  Case "13"
					 RptItem = "THR/Bypass"
                     If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV THR/Bypass", "Off"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV THR/Bypass", "On"
					 End If


				  Case "14"
					 RptItem = "RED DTS"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Red DTS", "Release"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Red DTS", "Wait"
					 End If
				  
				  Case "15"
					 RptItem = "Tuner Tuned"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Tuned", "NOT Tuned"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Tuned", "Tuned"
					 End If

				  Case "16"
					 RptItem = "DTS 96/24"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV DTS 96/24", "Off"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV DTS 96/24", "On"
					 End If
				  Case "20"	    
					 RptItem = "Power"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					 ElseIf RptStatusNbr = "03" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					 ElseIf RptStatusNbr = "04" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					 ElseIf RptStatusNbr = "05" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "On"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					 ElseIf RptStatusNbr = "06" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "On"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "Off"
					 ElseIf RptStatusNbr = "07" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Power Master", "Off"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 2", "Off"
						SetPropertyValue "Yamaha V2600 Settings.AV Power Zone 3", "On"
					 End If
					 
					 
				  Case "21"	    
					 RptItem = "Input"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "PHONO"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "CD"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "TUNER"
					 ElseIf RptStatusNbr = "03" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "CD-R"
					 ElseIf RptStatusNbr = "04" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "MD/TAPE"	  
					 ElseIf RptStatusNbr = "05" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "DVD"	  
					 ElseIf RptStatusNbr = "06" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "DTV"	  
					 ElseIf RptStatusNbr = "07" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "CBL/SAT"	  
					 ElseIf RptStatusNbr = "08" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "SAT"
					 ElseIf RptStatusNbr = "09" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "VCR1"			
					 ElseIf RptStatusNbr = "0A" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "DVR/VCR2"			
					 ElseIf RptStatusNbr = "0B" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "VCR3/DVR"			
					 ElseIf RptStatusNbr = "0C" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "V-AUX"			
					 ElseIf RptStatusNbr = "0E" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Input", "XM"			
					 End If
				  Case "22"	    
					 RptItem = "Audio Select"
					 If Mid(RptStatusNbr,2,1) = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Audio Select", "AUTO"
					 ElseIf Mid(RptStatusNbr,2,1) = "3" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Audio Select", "COAX/OPT"
					 ElseIf Mid(RptStatusNbr,2,1) = "4" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Audio Select", "ANALOG"
					 ElseIf Mid(RptStatusNbr,2,1) = "5" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Audio Select", "ANALOG ONLY"
					 ElseIf Mid(RptStatusNbr,2,1) = "8" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Audio Select", "HDMI"
					 End If
					 If Mid(RptStatusNbr,1,1) = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Decoder Mode", "AUTO"
					 ElseIf Mid(RptStatusNbr,1,1) = "1" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Decoder Mode", "DTS"
					 ElseIf Mid(RptStatusNbr,1,1) = "0" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Decoder Mode", "AAC"
					 End If

				  Case "23"	    
				     
					 RptItem = "Audio Mute"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Audio Mute", "Off"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Audio Mute", "On"
					 End If
				  Case "24"	    
					 RptItem = "Zone2 Input"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "PHONO"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "CD"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "TUNER"
					 ElseIf RptStatusNbr = "03" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "CD-R"
					 ElseIf RptStatusNbr = "04" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "MD/TAPE"	  
					 ElseIf RptStatusNbr = "05" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "DVD"	  
					 ElseIf RptStatusNbr = "06" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "DTV"	  
					 ElseIf RptStatusNbr = "07" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "CBL/SAT"	  
					 ElseIf RptStatusNbr = "08" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "SAT"
					 ElseIf RptStatusNbr = "09" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "VCR1"			
					 ElseIf RptStatusNbr = "0A" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "DVR/VCR2"			
					 ElseIf RptStatusNbr = "0C" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "V-AUX"			
					 ElseIf RptStatusNbr = "0E" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Input", "XM"			
					 End If

				 Case "25"	    
					 RptItem = "Zone2 Mute"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Mute", "Off"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Mute", "On"
					 End If
				  Case "26"	    
					 RptItem = "Main Volume"
                     	SetPropertyValue "Yamaha V2600 Settings.AV Master Volume", RptStatusNbr
						'sleep sleepvar
						ConvertToLinear(1)
				  Case "27"	    
					 RptItem = "Zone 2 Volume"
                     SetPropertyValue "Yamaha V2600 Settings.AV Zone 2 Volume", RptStatusNbr
					 'sleep sleepvar
				     ConvertToLinear(2)
 			      Case "28"	    
					 RptItem = "Program"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Hall in Munich"
  					    SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Hall B"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Hall C"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "04" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Hall C"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "05" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Hall in Vienna"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "06" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Live Concert"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "08" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Tokyo"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "09" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Freiburg"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "0A" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Royamount"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "0D" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Village Vanguard"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "0E" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "The Bottom Line"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "10" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "The Roxy Thtr"	
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "11" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Warehouse Loft"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
                     ElseIf RptStatusNbr = "12" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Arena"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "14" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Disco"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Entertainment"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "15" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Party"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "16" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Game"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Entertainment"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "On"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "17" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "7 Ch Stereo"	  
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Stereo"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "18" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Pop Rock"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "19" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "DJ"	  
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "1C" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Classical Opera"	  
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "1D" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Pavillion"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Music"
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"  
						SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"     
					 ElseIf RptStatusNbr = "20" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Mono Movie"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Entertainment"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "On"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "21" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "TV Sports"			
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Entertainment"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "On"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "24" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Spectacle"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Movie Theater"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "On"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "25" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Sci-Fi"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Movie Theater"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "On"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "28" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Adventure"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Movie Theater"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "On"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
                     ElseIf RptStatusNbr = "29" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "General"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Movie Theater"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "On"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
					 ElseIf RptStatusNbr = "2C" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Standard"	  
					 ElseIf RptStatusNbr = "2D" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Enhanced"
					 ElseIf RptStatusNbr = "30" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "PLII Movie"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Entertainment"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "On"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "On"
						
					 ElseIf RptStatusNbr = "31" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "PLII Music"	  
					 ElseIf RptStatusNbr = "32" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Neo:6 Movie"
					 ElseIf RptStatusNbr = "33" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "Neo:6 Music"			
					 ElseIf RptStatusNbr = "34" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "2 Ch Stereo"
						SetPropertyValue "Yamaha V2600 Settings.AV DSP Program Category", "Stereo"
	                    SetPropertyValue "Yamaha V2600 Settings.AV Indicator Cinema DSP", "Off"
                        SetPropertyValue "Yamaha V2600 Settings.AV Indicator HiFi DSP", "Off"
					 ElseIf RptStatusNbr = "35" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "STREO B 2CH Direct Stereo"			
					 ElseIf RptStatusNbr = "36" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "THX Cinema"
                     ElseIf RptStatusNbr = "37" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "THX Music"
					 ElseIf RptStatusNbr = "3C" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Program", "THX Game"
					 End If
					
  			      Case "29"	    
					 RptItem = "Tuner Page"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Page", "A"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Page", "B"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Page", "C"
					 ElseIf RptStatusNbr = "03" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Page", "D"
					 ElseIf RptStatusNbr = "04" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Page", "E"	  
					 End If
				  Case "2A"	    
					 RptItem = "Preset"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "1"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "2"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "3"
					 ElseIf RptStatusNbr = "03" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "4"
					 ElseIf RptStatusNbr = "04" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "5"	  
                     ElseIf RptStatusNbr = "05" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "6"
					 ElseIf RptStatusNbr = "06" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "7"
					 ElseIf RptStatusNbr = "07" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Preset Number", "8"	  
					 End If
				  Case "2B"	    
					 RptItem = "OSD"
				  Case "2C"	    
					 RptItem = "Sleep"	  
				  Case "2D"	    
					 RptItem = "EXTD SUR"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "Off"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "EX/ES"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "Discrete On"
					 ElseIf RptStatusNbr = "03" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "Auto"
					 ElseIf RptStatusNbr = "04" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "EX"	  
					 ElseIf RptStatusNbr = "05" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "PLIIx Movie"	  
					 ElseIf RptStatusNbr = "06" Then
						SetPropertyValue "Yamaha V2600 Settings.AV EXTD SUR", "PLIIx Music"
					 End If
				  Case "2E"	    
					 RptItem = "SP Relay A"
				  Case "2F"	    
					 RptItem = "SP Relay B"
				  Case "30"	    
					 RptItem = "System Memory Load"
				  Case "31"	    
					 RptItem = "System Memory Save"
				  Case "32"	    
					 RptItem = "Volume Memory Load"
				  Case "33"	    
					 RptItem = "Volume Memory Save"
				  Case "34"	    
					 RptItem = "Headphone"
				  Case "35"	    
					 RptItem = "Tuner Band"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Band", "FM"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Tuner Band", "AM"
					 End If
				  Case "36"	    
					 RptItem = "DC Trigger Output"
				  Case "37"	    
					 RptItem = "Zone2 Vol Memory Load"
				  Case "38"	    
					 RptItem = "Zone2 Vol Memory Save"
				  Case "39"	    
					 RptItem = "Dual Mono"
				  Case "3A"	    
					 RptItem = "DC1 Trigger Control"
				  Case "3B"	    
					 RptItem = "DC2 Trigger Control"
				  Case "3C"	    
					 RptItem = "DC2 Trigger Output"
				  Case "3D"	    
					 RptItem = "SP B SET"
				  Case "3E"	    
					 RptItem = "Zone2 Amp"
				  Case "3F"	    
					 RptItem = "Zone3 Amp"
				  Case "40"	    
					 RptItem = "Level Front R"
				  Case "41"	    
					 RptItem = "Level Front L"
				  Case "42"	    
					 RptItem = "Level Center"
				  Case "43"	    
					 RptItem = "Level Surround R"
				  Case "44"	    
					 RptItem = "Level Surround L"
				  Case "45"	    
					 RptItem = "Level Surround Back R"
				  Case "46"	    
					 RptItem = "Level Surround Back L"
				  Case "47"	    
					 RptItem = "Level Presence R"
				  Case "48"	    
					 RptItem = "Level Presence L"
				  Case "49"	    
					 RptItem = "Level Sub Woofer"
				  Case "4B"	    
					 RptItem = "Zone2 Control Bass"
				  Case "4C"	    
					 RptItem = "Zone2 Control Treble"
				  Case "4D"	    
					 RptItem = "Zone3 Control Bass"
				  Case "4E"	    
					 RptItem = "Zone3 Control Treble"
				  Case "50"	    
					 RptItem = "Main L/R Balance"
				  Case "51"	    
					 RptItem = "LFE Level SP"
				  Case "52"	    
					 RptItem = "LFE Level HP"
				  Case "53"	    
					 RptItem = "Audio Delay"
				  Case "58"	    
					 RptItem = "Wall paper"
				  Case "5A"	    
					 RptItem = "Max Volume"
				  Case "5B"	    
					 RptItem = "Initial Volume"
				  Case "5E"	    
					 RptItem = "Decoder Mode"
				  Case "5F"	    
					 RptItem = "Decoder Mode Set"
				  Case "60"	    
					 RptItem = "Audio Select"
				  Case "61"	    
					 RptItem = "Dimmer"
				  Case "62"	    
					 RptItem = "OSD Shift/GUI Position"
				  Case "63"	    
					 RptItem = "Gray Back"
				  Case "64"	    
					 RptItem = "Dymamic Range SP"
				  Case "65"	    
					 RptItem = "Dymamic Range HP"
				  Case "66"	    
					 RptItem = "Zone 2 Volume Out"
				  Case "68"	    
					 RptItem = "Memory Guard"
				  Case "69"	    
					 RptItem = "Video Conversion"
				  Case "6B"	    
					 RptItem = "Zone3 Volume Out"
				  Case "6C"	    
					 RptItem = "Zone 2 OSD"
				  Case "6E"	    
					 RptItem = "2Ch Decoder"
				  Case "70"	    
					 RptItem = "Center SP"
				  Case "71"	    
					 RptItem = "Front"
				  Case "72"	    
					 RptItem = "Surround SP"
				  Case "73"	    
					 RptItem = "Surround Back"
				  Case "74"	    
					 RptItem = "Presence"
				  Case "75"	    
					 RptItem = "LFE Bass Out"
				  Case "76"	    
					 RptItem = "Subwoofer Out"
				  Case "7B"	    
					 RptItem = "Multi CH Select"
				  Case "7D"	    
					 RptItem = "PR/SB Priority"
				  Case "7E"	    
					 RptItem = "Subwoofer Crossover"
				  Case "80"	    
					 RptItem = "Test"
				  Case "85"	    
					 RptItem = "Component I/P"
				  Case "86"	    
					 RptItem = "HDMI I/P"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV HDMI I/P", "Off"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV HDMI I/P", "Off"
					 End If
				  Case "87"	    
					 RptItem = "HDMI Up-Scaling"
				  Case "88"	    
					 RptItem = "HDMI Aspect"
				  Case "8A"	    
					 RptItem = "THX SB Dist"
				  Case "8B"	    
					 RptItem = "Night Mode Parameter"
				  Case "8C"	    
					 RptItem = "Pure Direct"
				  Case "8F"	    
					 RptItem = "HDMI Support Audio"
				  Case "90"	    
					 RptItem = "XM Preset Page"
				  Case "91"	    
					 RptItem = "XM Preset Number"
				  Case "92"	    
					 RptItem = "XM Search Mode"
				  Case "93"	    
					 RptItem = "XM Display Time"
				  Case "94"	    
					 RptItem = "XM CH Number"
				  Case "95"	    
					 RptItem = "XM Display Hold/Release"
				  Case "A0"	    
					 RptItem = "Zone 3 Input"
					 If RptStatusNbr = "00" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "PHONO"
					 ElseIf RptStatusNbr = "01" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "CD"
					 ElseIf RptStatusNbr = "02" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "TUNER"
					 ElseIf RptStatusNbr = "03" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "CD-R"
					 ElseIf RptStatusNbr = "04" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "MD/TAPE"	  
					 ElseIf RptStatusNbr = "05" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "DVD"	  
					 ElseIf RptStatusNbr = "06" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "DTV"	  
					 ElseIf RptStatusNbr = "07" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "CBL/SAT"	  
					 ElseIf RptStatusNbr = "08" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "SAT"
					 ElseIf RptStatusNbr = "09" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "VCR1"			
					 ElseIf RptStatusNbr = "0A" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "DVR/VCR2"			
					 ElseIf RptStatusNbr = "0C" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "V-AUX"			
					 ElseIf RptStatusNbr = "0E" Then
						SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Input", "XM"			
					 End If

				  Case "A1"	    
					 RptItem = "Zone 3 Mute"
				  Case "A2"	    
					 RptItem = "Zone 3 Volume"
                     SetPropertyValue "Yamaha V2600 Settings.AV Zone 3 Volume", RptStatusNbr
					 'sleep sleepvar
				     ConvertToLinear(3)
				  Case "A3"	    
					 RptItem = "Zone 3 Volume Memory Load"
				  Case "A4"	    
					 RptItem = "Zone 3 Volume Memory Save"
				  Case "A5"	    
					 RptItem = "Mute Type"
				  Case "A7"	    
					 RptItem = "EQ Select Type"
				  Case "A8"	    
					 RptItem = "Tone Bypass"
				  Case "B0"	    
					 RptItem = "Advanced Setup"
				  Case "B1"	    
					 RptItem = "Remote ID for AMP"
				  Case "B2"	    
					 RptItem = "Fan Control"
				  Case "B3"	    
					 RptItem = "Speaker Impedence"
				  Case "B4"	    
					 RptItem = "Tuner Step"
				  Case "B5"	    
					 RptItem = "Remote ID for Tuner"
				  Case "B6"	    
					 RptItem = "Language"
				  Case "B7"	    
					 RptItem = "User Preset"
				  Case "B8"	    
					 RptItem = "Video Reset"
				  Case "B9"	    
					 RptItem = "Remote Sensor"
				  Case "BA"	    
					 RptItem = "Remote ID for XM"
				  Case "BB"	    
					 RptItem = "Bi-AMP"
				  Case "BC"	    
					 RptItem = "TV Format"
				  Case "BD"	    
					 RptItem = "Wake on RS232"
			   End Select
			  'MsgBox  RptItem
        End Select
   
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
