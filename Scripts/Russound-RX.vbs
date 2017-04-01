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
   	
   	Select Case Left(DataItem, 1)
	
		    ' Configuration Command DC2
		Case Chr(18)


		End Select
End Sub