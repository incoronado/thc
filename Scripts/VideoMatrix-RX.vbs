Dim Data, SleepVar
SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))

Sleep SleepVar
ReadSerialData GetPropertyValue ("HDMI Matrix.Received Data")


Sub ReadSerialData(Data)
	If Mid(Data,2,1) = Chr(97) Then
		SetpropertyValue "HDMI Matrix Settings.Debug", "New Command"
		If  Mid(Data,12,1) = Chr(01) Then
			SetpropertyValue "HDMI Matrix Settings.HDMI Power State", "On"
		Else
			SetpropertyValue "HDMI Matrix Settings.HDMI Power State", "Off"
		End if
		
		Select Case Mid(Data,7,1) 
			Case Chr(01)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone A", "1"
			Case Chr(02)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone A", "2"
			Case Chr(04) 
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone A", "3"
			Case Chr(08) 	
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone A", "4"
		End Select
		Sleep SleepVar
		SetpropertyValue "System.Zone A Selected Source", GetPropertyValue("System.Matrix Source " & GetPropertyValue("HDMI Matrix Settings.Matrix Zone A") & " Name")
		Select Case Mid(Data,8,1) 
			Case Chr(01)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone B", "1"
			Case Chr(02)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone B", "2"
			Case Chr(04) 
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone B", "3"
			Case Chr(08) 	
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone B", "4"
		End Select 
		Sleep SleepVar
		SetpropertyValue "System.Zone B Selected Source", GetPropertyValue("System.Matrix Source " & GetPropertyValue("HDMI Matrix Settings.Matrix Zone B") & " Name")	
		
		Select Case Mid(Data,9,1) 
			Case Chr(01)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone C", "1"
			Case Chr(02)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone C", "2"
			Case Chr(04) 
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone C", "3"
			Case Chr(08) 	
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone C", "4"
		End Select 
		Sleep SleepVar
		SetpropertyValue "System.Zone C Selected Source", GetPropertyValue("System.Matrix Source " & GetPropertyValue("HDMI Matrix Settings.Matrix Zone C") & " Name")
		
		Select Case Mid(Data,10,1) 
			Case Chr(01)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone D", "1"
			Case Chr(02)
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone D", "2"
			Case Chr(04) 
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone D", "3"
			Case Chr(08) 	
				SetpropertyValue "HDMI Matrix Settings.Matrix Zone D", "4"
		End Select 
		Sleep SleepVar
		SetpropertyValue "System.Zone D Selected Source", GetPropertyValue("System.Matrix Source " & GetPropertyValue("HDMI Matrix Settings.Matrix Zone D") & " Name")			
	End If	
End Sub
