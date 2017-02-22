Option Explicit
'On Error Resume Next

Dim DataItem, SleepVar
DataItem = GetPropertyValue ("USBUIRT.IR Code Number Received")
SetModeState "Block IR", "Inactive" 
SetPropertyValue "IR Block.Running", "Yes"

Select Case DataItem

	Case "133"
		SendSubscriberMessage 9, "System.BedroomIR.10.VolumeUp"
	Case "142"
		SendSubscriberMessage 9, "System.BedroomIR.10.VolumeUp"
	Case "135"
		SendSubscriberMessage 9, "System.BedroomIR.10.VolumeDown"
	Case "141"
		SendSubscriberMessage 9, "System.BedroomIR.10.VolumeDown"
	Case "3"
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:CableTVOn:A"
	Case "1"	
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:AppleTVOn:A"
	Case "4"	
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:BlueRayOn:A"
	Case "5"
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:XBMCOn:A"
	Case "48"
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:AllMediaOff:A"
	Case "147"	
		SendSubscriberMessage 9, "System.IRRemote.10.ToggleBedroomTV"
		SendSubscriberMessage 9, "MRA.GalaxyTabA1.10.off:5"
	Case "132"	
		SendSubscriberMessage 9, "System.IRRemote.10.ToggleBedroomTV"
		SendSubscriberMessage 9, "MRA.GalaxyTabA1.10.off:5"
	Case "127"	
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:CableTVOn:B"
		SendSubscriberMessage 9, "System.GalaxyTabA1.10.WestinghouseTVPower"
		
	Case "143"	
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:CableTVOn:B"
		SendSubscriberMessage 9, "System.GalaxyTabA1.10.WestinghouseTVPower"
	Case "128"	
		'System.GalaxyTabA1.10.WestinghouseTVPower
		SendSubscriberMessage 9, "System.GalaxyTabA1.10.AllMediaOff"		
	Case "130"
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:AppleTVOn:B"
	Case "144"
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:AppleTVOn:B"
	Case "129"
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:XBMCOn:B"
	Case "148"
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:XBMCOn:B"
	Case "145"
		SendSubscriberMessage 9, "System.GalaxyTabA1.10.XBMC:Guide"
	Case "146"
		SendSubscriberMessage 9, "System.GalaxyTabA1.10.XBMC:Guide"
	Case "131"
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:BlueRayOn:B"
	Case "125"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:Previous"	
	Case "126"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:Previous"		
	Case "123"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:Next"
	Case "124"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:Next"
	Case "120"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:RW"
	Case "121"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:RW"
	Case "113"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:FF"
	Case "112"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:FF"
	Case "118"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:Stop"
	Case "119"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:Stop"
	Case "114"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:Play"
	Case "115"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:Play"
	Case "116"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:Pause"
	Case "117"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Player:Pause"
	Case "60"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Select"
	Case "108"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Select"
	Case "106"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Back"
	Case "109"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Back"
	Case "107"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Home"
	Case "110"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Home"
	Case "59"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Left"
	Case "105"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Left"
	Case "58"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Right"
	Case "104"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Right"
	Case "57"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Down"
	Case "64"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Down"		
	Case "56"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Up"
	Case "65"
		SendSubscriberMessage 9, "System.IRRemote.10.XBMC:Input:Up"
	Case "94"
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:XBMCOn:A"
	Case "111"
		SendSubscriberMessage 9, "System.IRRemote.10.AVOn:XBMCOn:A"
		
	

	
End Select	





' ********************** Global Functions: Needs to Be in all Subscriber Files****************************
Sub SendSubscriberMessage(SubscriberID, SubscriberMessage)
	do 
		If GetPropertyValue("Subscriber-" & CStr(SubscriberID) & ".DispatchMessage") = "Idle" Then
			'SetPropertyValue "System.Debug", "Got Subscriber"
			'Perform A Crude Form of Record Locking 
			BlockSubscriber(SubscriberID)	
			SetpropertyValue "Subscriber-" & Cstr(SubscriberID) & ".DispatchMessage", SubscriberMessage
			UnBlockSubscriber(SubscriberID) 
			Exit Do
		Else
			Sleep GetRandomNumber (1,100)
		End if
	Loop

End Sub

Sub BlockSubscriber(SubscriberID) 
	SetModeState "Subscriber-" & CStr(SubscriberID), "Inactive"
End Sub	

Sub UnBlockSubscriber(SubscriberID) 
	SetModeState "Subscriber-" & CStr(SubscriberID), "Active"
End Sub	

Function GetRandomNumber (min,max)
	Randomize
	GetRandomNumber = Int((max-min+1)*Rnd+min)
End Function

Sub HBRemoteList
	Dim NoMoreRemotes, RemoteCount, HBRList
	HBRList = ""
	RemoteCount = 0
	NoMoreRemotes = 0
	Do Until NoMoreRemotes = 1
		If ((GetPropertyValue ("Remote-" + CStr(RemoteCount+1) + ".Remote Name") <> "* error *") And (GetPropertyValue ("Remote-" + CStr(RemoteCount+1) + ".Remote Type") = "HB")) Then
			HBRList = HBRList & CStr(RemoteCount+1) & "," 	
			RemoteCount=RemoteCount + 1
		Else
			NoMoreRemotes = 1
		End if
		
	Loop
	SetPropertyValue "System.HBRemoteCount", CInt(RemoteCount)
	SetPropertyValue "System.HBRemoteList", Left(HBRList,Len(HBRList)-1) 
End Sub

' EOF ********************** Needs to Be in all Subscriber Files****************************