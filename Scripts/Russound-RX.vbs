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
	If ValidateRNETChecksum(Data) = True Then
		'SetPropertyValue "Multiroom Audio Settings.Debug 2", "Good"
   	Else
   		'SetPropertyValue "Multiroom Audio Settings.Debug 2", "Bad"
   	End If

   	Select Case Left(DataItem, 1)
	
		    ' Configuration Command DC2
		Case Chr(18)


		End Select
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
	SetPropertyValue "Multiroom Audio Settings.Debug 2", MsgChecksum 
	If ComputeRNETChecksum(hexstr) = MsgChecksum then
		ValidateRNETChecksum = True
	Else
		ValidateRNETChecksum = False
	End If	
End Function