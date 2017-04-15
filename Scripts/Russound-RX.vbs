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
	Dim DataItem, HexBytes, ComputedChecksum, i, MessageLength, SourceNo, MessageStr, RNETHexBytes, x, RNETPayloadHeader
	'SetPropertyValue "Yamaha V2600 Settings.AV Debug", data
	MessageStr = ""
	
	

	If ValidateRNETChecksum(Data) = False Then
		  
   	Else

   		HexBytes=split(Data," ")

   		Select Case HexBytes(3)
		    ' Read Source Broadcast Display Feedback		    
			Case "79"

				RNETHexBytes=split(RNETMessage(Data)," ")
				'Overall Payload Size
				MessageLength = CLng("&h" & HexBytes(18))
				SourceNo = CLng(Right(RNETHexBytes(1),1)) + 1
				SetPropertyValue "Multiroom Audio Settings.Debug", RNETMessage(Data)
				If RNETHexBytes(0) = "23" Then
					For i = 23 To MessageLength + 19
						'MessageStr = MessageStr & chr(CLng("&h" & HexBytes(i)))
					Next
					SetPropertyValue "Multiroom Audio Settings.Debug", MessageStr		
				ElseIf  RNETHexBytes(0) = "60" Then
				    RNETPayloadHeader = ""
				    For i = 0 To ubound(RNETHexBytes)

						If RNETHexBytes(i) = "1C" Then
						    for x = CLng("&h" & HexBytes(i+1)) To ubound(RNETHexBytes)
								MessageStr = MessageStr + chr(CLng("&h" & RNETHexBytes(x)))
							Next	
							Exit For
						End If	

					Next	

					Select Case RNETHexbytes(4)
						Case "06"
							SetPropertyValue "Russound.XM Format", MessageStr
						Case "02"	
							SetPropertyValue "Russound.XM Artist", MessageStr
						Case "01"	
							SetPropertyValue "Russound.XM Song", MessageStr	
						Case "04"	
							SetPropertyValue "Russound.XM Album", MessageStr
						Case "15"	
							SetPropertyValue "Russound.XM URL", MessageStr
						Case "0F"	
							SetPropertyValue "Russound.Streamer Source", MessageStr	
						Case "07"
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
	Dim MessageLength, MessageStr, HexBytes, i, objLogFile, HousebotLocation, ObjFSO
	HousebotLocation = "C:\Program Files (x86)\housebot\"

	HexBytes=split(Data," ")
	'Overall Payload Size
	MessageLength = CLng("&h" & HexBytes(18))
		For i = 20 To ubound(HexBytes) - 1
			MessageStr = MessageStr & HexBytes(i) + " "
		Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objLogFile = objFSO.OpenTextFile(HousebotLocation & "config\scripts\RNETMessage.log", 8, true)
	objLogFile.Write Trim(MessageStr) & vbCrLf
	objLogFile.Close

	RNETMessage =  Trim(MessageStr)
End Function