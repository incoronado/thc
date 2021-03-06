Option Explicit
'On Error Resume Next

Dim DataItem, SleepVar
DataItem = GetPropertyValue ("USBUIRT.IR Code Number Received")
SetModeState "Block IR", "Inactive" 
SetPropertyValue "IR Block.Running", "Yes"

Select Case DataItem

	Case "133"
		SendSubscriberMessage 10, "Russound.GalaxyTabA1.10.Volume:Master Bedroom:Up"
	Case "142"
		SendSubscriberMessage 10, "Russound.GalaxyTabA1.10.Volume:Master Bedroom:Up"
	Case "135"
		SendSubscriberMessage 10, "Russound.GalaxyTabA1.10.Volume:Master Bedroom:Down"
	Case "141"
		SendSubscriberMessage 10, "Russound.GalaxyTabA1.10.Volume:Master Bedroom:Down"
	Case "3"
		'SendSubscriberMessage 10, "System.IRRemote.10.AVOn:CableTVOn:A"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Cable TV:Living Room"
	Case "1"	
		'SendSubscriberMessage 10, "System.IRRemote.10.AVOn:AppleTVOn:A"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Apple TV:Living Room"
	Case "4"	
		'SendSubscriberMessage 10, "System.IRRemote.10.AVOn:BlueRayOn:A"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Blu-Ray:Living Room"
	Case "5"
		'SendSubscriberMessage 10, "System.IRRemote.10.AVOn:XBMCOn:A"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Kodi:Living Room"
	Case "48"
		'SendSubscriberMessage 10, "System.IRRemote.10.AVOn:AllMediaOff:A"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOff2:Living Room"	
	Case "147"	
		SendSubscriberMessage 10, "System.IRRemote.10.AVOff2:Master Bedroom"
	Case "132"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOff2:Master Bedroom"	
	Case "127"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Cable TV:Master Bedroom"
	Case "143"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Cable TV:Master Bedroom"
	Case "128"	
		SendSubscriberMessage 10, "System.GalaxyTabA1.10.AllMediaOff"		
	Case "130"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Apple TV:Master Bedroom"
	Case "144"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Apple TV:Master Bedroom"
	Case "129"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Kodi:Master Bedroom"
	Case "148"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Kodi:Master Bedroom"
	Case "145"
		SendSubscriberMessage 10, "System.GalaxyTabA1.10.XBMC:Guide"
	Case "146"
		SendSubscriberMessage 10, "System.GalaxyTabA1.10.XBMC:Guide"
	Case "131"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Blu-Ray:Master Bedroom"
	Case "171"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn2:Blu-Ray:Master Bedroom"	
	Case "125"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:Previous"	
	Case "126"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:Previous"		
	Case "123"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:Next"
	Case "124"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:Next"
	Case "120"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:RW"
	Case "121"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:RW"
	Case "113"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:FF"
	Case "112"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:FF"
	Case "118"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:Stop"
	Case "119"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:Stop"
	Case "114"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:Play"
	Case "115"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:Play"
	Case "116"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:Pause"
	Case "117"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Player:Pause"
	Case "60"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Select"
	Case "108"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Select"
	Case "106"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Back"
	Case "109"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Back"
	Case "107"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Home"
	Case "110"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Home"
	Case "59"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Left"
	Case "105"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Left"
	Case "58"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Right"
	Case "104"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Right"
	Case "57"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Down"
	Case "64"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Down"		
	Case "56"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Up"
	Case "65"
		SendSubscriberMessage 10, "System.IRRemote.10.XBMC:Input:Up"
	Case "94"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn:XBMCOn:A"
	Case "111"
		SendSubscriberMessage 10, "System.IRRemote.10.AVOn:XBMCOn:A"
		
	

	
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