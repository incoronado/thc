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


Dim Data, SleepVar
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))
ReadSerialData GetPropertyValue("Multiroom Audio Amplifier.Received Hex Data")
Sleep SleepVar

Sub ReadSerialData(Data)
	Dim DataItem, HexBytes, ComputedChecksum, i, MessageLength, SourceNo, MessageStr
	'SetPropertyValue "Yamaha V2600 Settings.AV Debug", data
	MessageStr = ""
	
	

	If ValidateRNETChecksum(Data) = False Then
		  
   	Else
   		HexBytes=split(Data," ")
   		Select Case HexBytes(3)
		    ' Read Source Broadcast Display Feedback
			Case "79"
				'Overall Payload Size
				MessageLength = CLng("&h" & HexBytes(18))
				SourceNo = CLng(Right(HexBytes(20),1)) + 1
				SetPropertyValue "Multiroom Audio Settings.Debug", RNETMessage(Data)
				If HexBytes(20) = "234" Then
					For i = 23 To MessageLength + 19
						'MessageStr = MessageStr & chr(CLng("&h" & HexBytes(i)))
					Next
					SetPropertyValue "Multiroom Audio Settings.Debug", MessageStr		
				ElseIf HexBytes(20) = "604" Then	
					Select Case Hexbytes(24)
						Case "06"
							For i = 29 To CLng("&h" & HexBytes(28)) + 28
								MessageStr = MessageStr & chr(CLng("&h" & HexBytes(i)))
							Next	
							SetPropertyValue "Russound.XM Format", MessageStr
						Case "02"	
							For i = 32 To CLng("&h" & HexBytes(31)) + 31
								MessageStr = MessageStr & chr(CLng("&h" & HexBytes(i)))
							Next	
							SetPropertyValue "Russound.XM Artist", MessageStr
						Case "01"	
							For i = 32 To CLng("&h" & HexBytes(31)) + 31
								MessageStr = MessageStr & chr(CLng("&h" & HexBytes(i)))
							Next	
							SetPropertyValue "Russound.XM Song", MessageStr	
						Case "04"	
							For i = 32 To CLng("&h" & HexBytes(31)) + 31
								MessageStr = MessageStr & chr(CLng("&h" & HexBytes(i)))
							Next	
							SetPropertyValue "Russound.XM Album", MessageStr
						Case "15"	
							For i = 32 To CLng("&h" & HexBytes(31)) + 31
								MessageStr = MessageStr & chr(CLng("&h" & HexBytes(i)))
							Next	
							SetPropertyValue "Russound.XM URL", MessageStr
						Case "0F"	
							For i = 32 To CLng("&h" & HexBytes(31)) + 31
								MessageStr = MessageStr & chr(CLng("&h" & HexBytes(i)))
							Next	
							SetPropertyValue "Russound.Streamer Source", MessageStr	
						Case "07"
							For i = 29 To CLng("&h" & HexBytes(28)) + 28
								MessageStr = MessageStr & chr(CLng("&h" & HexBytes(i)))
							Next	
							SetPropertyValue "Russound.XM Station", MessageStr				
					End Select
							


				End if	
				

				
		End Select
   	End If
End Sub


Function ComputeRNETChecksum(hexstr)
	Dim HexBytes, ComputedChecksum, i
	HexBytes=split(hexstr," ")
	For i = 0 To ubound(HexBytes) - 1
		'Add the HEX value of every byte in the message that precedes the Checksum in decimal
		ComputedChecksum = ComputedChecksum + CLng("&h" & HexBytes(i))
	Next
	'Count the number of bytes which precede the Checksum still maintaining decimal
	ComputedChecksum = ComputedChecksum + ubound(HexBytes) 
	' This value is then AND-ed with the HEX value 0x007F or 127 Dec
	ComputedChecksum = ComputedChecksum And 127
	ComputeRNETChecksum=cStr(hex(ComputedChecksum))
End Function

Function ValidateRNETChecksum(hexstr)
	Dim HexBytes, MsgChecksum
	HexBytes=split(hexstr," ")
	MsgChecksum = HexBytes(ubound(HexBytes))
	'SetPropertyValue "Multiroom Audio Settings.Debug 2", MsgChecksum 
	If ComputeRNETChecksum(hexstr) = MsgChecksum then
		ValidateRNETChecksum = True
	Else
		ValidateRNETChecksum = False
	End If	
End Function

Function RNETMessage(Data)
	Dim MessageLength, MessageStr, HexBytes, i
	HexBytes=split(Data," ")
	'Overall Payload Size
	MessageLength = CLng("&h" & HexBytes(18))
		For i = 20 To ubound(HexBytes) - 1
			MessageStr = MessageStr & HexBytes(i) + " "
		Next
	RNETMessage =  Trim(MessageStr)
End Function