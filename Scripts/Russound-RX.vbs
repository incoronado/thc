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
	
	Dim RptStatusNbr, RptCmdTyp, RptCmdNbr, RptItem, DataItem, DataBuffer, msglength
	'SetPropertyValue "Yamaha V2600 Settings.AV Debug", data
	DataItem = Data
	SetPropertyValue "Multiroom Audio Settings.Debug", Data 
	SetPropertyValue "Multiroom Audio Settings.Debug 2", RNETChecksum(Data)
   	
   	Select Case Left(DataItem, 1)
	
		    ' Configuration Command DC2
		Case Chr(18)


		End Select
End Sub


Function RNETChecksum(hexstr)
	Dim HexBytes, ComputedChecksum, i
	HexBytes=split(hexstr," ")
	
	'F0 7D 00 7F 00 00 7F 05 02 01 00 02 01 00 66 01 00 00 80 02 01 75'
	
	For i = 0 To ubound(HexBytes) - 2
		'Add the HEX value of every byte in the message that precedes the Checksum in decimal'
		ComputedChecksum = ComputedChecksum + CLng("&h" & i)
	Next
	'Count the number of bytes which precede the Checksum still maintaining decimal
	ComputedChecksum = ComputedChecksum + ubound(HexBytes) - 1
	' This value is then AND-ed with the HEX value 0x007F or 127 Dec
	ComputedChecksum = ComputedChecksum And 127

	RNETChecksum=cStr(hex(ComputedChecksum))
End Function
