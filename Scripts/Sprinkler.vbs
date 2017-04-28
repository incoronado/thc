Option Explicit
On Error Resume Next
Dim Action, Status, SleepVar


SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))

Do
	Sleep SleepVar
	Action = GetPropertyValue ("Sprinklers.Action")
	If Action <> "Idle" Then
		Message_Handler(Action)
		SetpropertyValue "Sprinklers.Action", "Idle"		
	End If

Loop


Sub Message_Handler(Message)
Dim a, b
	a=split(Message,".")
	Select Case a(0)
		'Sprinker Class Message Received
		'Sprinkler.GalaxyTabA1.10.ToggleZone:2
		Case "Sprinkler"
			b=split(a(3),":")
			select case b(0)
				Case "ToggleZone"			    : ToggleZone b(1)
				'i.e., Sprinkler.GalaxyTabA1.10.ToggleZone:2
				Case "ToggleDOW"		        : ToggleDOW b(1)
				'i.e., Sprinkler.GalaxyTabA1.10.ToggleDOW:Mon
				Case "ToggleMasterState"		: ToggleMasterState
				'i.e., Sprinkler.GalaxyTabA1.10.ToggleMasterState
				Case "ToggleSprinklersStatus"	: ToggleSprinklersStatus
				'i.e., Sprinkler.GalaxyTabA1.10.ToggleSprinklersStatus
				Case "RunSprinklerSchedule"	    : RunSprinklerSchedule
				'i.e., Sprinkler.GalaxyTabA1.10.RunSprinklerSchedule
				Case "RunNextZone"	            : RunNextZone
				'i.e., Sprinkler.GalaxyTabA1.10.RunNextZone
				Case "TimeDown"	            	: TimeDown b(1)
				Case "TimeUp"	            	: TimeUp b(1)
				Case "StartTimeUp"				: StartTimeUp b(1)
				Case "StartTimeDown"			: StartTimeDown b(1)
				Case "TurnAllSprinklersOff"    	: TurnAllSprinklersOff
				'i.e., Sprinkler.GalaxyTabA1.10.TurnAllSprinklersOff
				Case "SprinkerStartTime"		: sprinklerstarttime
				Case "Start Zone 1"				: StartZone, "1"
				Case "Start Zone 2"				: StartZone, "2"
				Case "Start Zone 3"				: StartZone, "3"
				Case "Start Zone 4"				: StartZone, "4"
				Case "Stop Zone 1"				: StopZone, "1"
				Case "Stop Zone 2"				: StopZone, "2"
				Case "Stop Zone 3"				: StopZone, "3"
				Case "Stop Zone 4"				: StopZone, "4"
				
			End Select 
		End Select
 End Sub

Sub StartZone(Zone)
   If GetPropertyValue ("Sprinklers.Sprinklers Zone " & Zone & " State") = "Stopped" And GetPropertyValue("Sprinklers.Sprinklers Zone " & Zone & " Time") > 0 Then
      CheckAllZonesOff
      SetpropertyValue "Sprinklers.Sprinklers Zone " & Zone & " State", "Running"
	  SetpropertyValue "Sprinkler System.Sprinklers Zone " & Zone & " State", "On"
	  SetpropertyValue "Sprinkler Zone Timer.Sleep Time", "Days=0, Hours=00, Minutes=" & GetPropertyValue("Sprinklers.Sprinklers Zone " & Zone & " Time") & ", Seconds=00, MilliSeconds=0"
	  SetpropertyValue "Sprinkler Zone Timer.Running", "Yes"
   End if 
End Sub

Sub StopZone(Zone)
	  SetpropertyValue "Sprinklers.Sprinklers Zone " & Zone & " State", "Stopped"
	  SetpropertyValue "Sprinkler System.Sprinklers Zone " & Zone & " State", "Off"
	  TurnAllSprinklersOff
End Sub

 
Sub ToggleZone(Zone)
   If GetPropertyValue ("Sprinklers.Sprinklers Zone " & Zone & " State") = "Stopped" And GetPropertyValue("Sprinklers.Sprinklers Zone " & Zone & " Time") > 0 Then
      CheckAllZonesOff
      SetpropertyValue "Sprinklers.Sprinklers Zone " & Zone & " State", "Running"
	  SetpropertyValue "Sprinkler System.Sprinklers Zone " & Zone & " State", "On"
	  SetpropertyValue "Sprinkler Zone Timer.Sleep Time", "Days=0, Hours=00, Minutes=" & GetPropertyValue("Sprinklers.Sprinklers Zone " & Zone & " Time") & ", Seconds=00, MilliSeconds=0"
	  SetpropertyValue "Sprinkler Zone Timer.Running", "Yes"
   Else 
      SetpropertyValue "Sprinklers.Sprinklers Zone " & Zone & " State", "Stopped"
	  SetpropertyValue "Sprinkler System.Sprinklers Zone " & Zone & " State", "Off"
	  TurnAllSprinklersOff
   End If
End Sub


Sub ToggleMasterState
   If GetPropertyValue ("Sprinklers.Sprinklers Master State") = "Off" Then
      SetpropertyValue "Sprinklers.Sprinklers Master State", "On"
   Else
      CheckAllZonesOff
      SetpropertyValue "Sprinklers.Sprinklers Master State", "Off"
   End If
End Sub

Sub CheckAllZonesOff
	Dim ZoneCount, i
	ZoneCount = CountSprinklerZones
	For i = 1 To ZoneCount
		If GetPropertyValue ("Sprinklers.Sprinklers Zone " & CStr(i) & " State") = "Running" Then
			SetpropertyValue "Sprinklers.Sprinklers Zone " & CStr(i) & " State", "Stopped"
		End if
		SetpropertyValue "Rain8.Rain 8 Command", "Zone" & CStr(i) & " Off"
		sleep 100	
	Next
End Sub

Sub ToggleDOW(DOW)
   If GetPropertyValue ("Sprinklers.Sprinklers DOW " & DOW) = "Off" Then
      SetpropertyValue "Sprinklers.Sprinklers DOW " & DOW, "On"
   Else 
      SetpropertyValue "Sprinklers.Sprinklers DOW " & DOW, "Off"
   End If
End Sub

Sub ToggleSprinklersStatus
   If GetPropertyValue ("Sprinklers.Sprinklers Status") = "Stopped" Then
      SetpropertyValue "Sprinklers.Sprinklers Status", "Running"
   Else 
      TurnAllSprinklersOff
      SetpropertyValue "Sprinklers.Sprinklers Status", "Stopped"
   End If
End Sub


Sub RunSprinklerSchedule
    If GetPropertyValue ("Sprinklers.Sprinklers Master State") = "On" Then
	   If GetPropertyValue ("Sprinklers.Sprinklers DOW " & ShortDOW(GetPropertyValue("System Time.Day Of Week"))) = "On" Then
          CheckAllZonesOff
          SetpropertyValue "Sprinklers.Sprinkler Running Zone", 1
          'Load Up Timer with Sprinkler Run Time
          SetpropertyValue "Sprinkler Schedule Timer.Sleep Time", "Days=0, Hours=00, Minutes=" & GetPropertyValue("Sprinklers.Sprinklers Zone 1 Time") & ", Seconds=00, MilliSeconds=0"
          SetpropertyValue "Sprinklers.Sprinklers Zone 1 State", "Running"
          SetpropertyValue "Sprinkler Schedule Timer.Running", "Yes"
          SetpropertyValue "Rain8.Rain 8 Command", "Zone1 On"
       Else
	      SetpropertyValue "Sprinklers.Sprinklers Status", "Stopped"
       End If
    Else
       SetpropertyValue "Sprinklers.Sprinklers Status", "Stopped"
    End If
End Sub

Sub RunNextZone
  Dim ZoneStr, ZoneCount
  ZoneCount = CInt(CountSprinklerZones)
  SetPropertyValue "Sprinklers.Debug", ZoneCount
  If CInt(GetPropertyValue("Sprinklers.Sprinkler Running Zone")) =>  ZoneCount Then
	 SetpropertyValue "Sprinklers.Sprinklers Zone " & GetPropertyValue("Sprinklers.Sprinkler Running Zone") & " State", "Stopped"
	 SetpropertyValue "Sprinklers.Sprinklers Status", "Stopped"
	 TurnAllSprinklersOff
  Else
     SetpropertyValue "Sprinklers.Sprinklers Zone " & GetPropertyValue("Sprinklers.Sprinkler Running Zone") & " State", "Stopped"
     SetpropertyValue "Rain8.Rain 8 Command", "Zone" & GetPropertyValue("Sprinklers.Sprinkler Running Zone") & " Off"
     SetpropertyValue "Sprinklers.Sprinkler Running Zone", GetPropertyValue("Sprinklers.Sprinkler Running Zone") + 1
     sleep 50
	 ZoneStr = "Sprinklers.Sprinklers Zone " & GetPropertyValue("Sprinklers.Sprinkler Running Zone") & " Time"
	 If GetPropertyValue(ZoneStr) = "00"  Then
		SetpropertyValue "Sprinkler Schedule Timer.Sleep Time", "Days=0, Hours=00, Minutes=00, Seconds=02, MilliSeconds=0"
	 Else
		SetpropertyValue "Sprinkler Schedule Timer.Sleep Time", "Days=0, Hours=00, Minutes=" & GetPropertyValue(ZoneStr) & ", Seconds=00, MilliSeconds=0"
	 End If
	 SetpropertyValue "Sprinklers.Sprinklers Zone " & GetPropertyValue("Sprinklers.Sprinkler Running Zone") & " State", "Running"
	 SetpropertyValue "Sprinkler Schedule Timer.Running", "Yes"
	 SetpropertyValue "Rain8.Rain 8 Command", "Zone" & GetPropertyValue("Sprinklers.Sprinkler Running Zone") & " On"
 
  End If

End Sub

Sub TurnAllSprinklersOff
      SetModeState "Block Sprinklers", "Inactive"
	  Sleep 200
      'SetpropertyValue "Sprinkler Zone Timer.Running", "No"
      SetpropertyValue "Sprinkler Schedule Timer.Running", "No"
	  Sleep 50
	  SetModeState "Block Sprinklers", "Active"
	  SetpropertyValue "Sprinklers.Sprinkler Running Zone", 0
	  CheckAllZonesOff
End Sub

Function ShortDOW(dow)
Dim dowAcronym
   If dow = "Sunday" Then
      dowAcronym = "Sun"
   Elseif dow = "Monday" Then
      dowAcronym = "Mon"
   Elseif dow = "Tuesday" Then
      dowAcronym = "Tue"
   Elseif dow = "Wednesday" Then
      dowAcronym = "Wed"
   Elseif dow = "Thursday" Then
      dowAcronym = "Thu"
   Elseif dow = "Friday" Then
      dowAcronym = "Fri"
   Elseif dow = "Saturday" Then
      dowAcronym = "Sat"
   End if
   ShortDOW = dowAcronym
End Function

Sub sprinklerstarttime
		SetPropertyValue "Sprinklers.Sprinkler Start Time Text", Trim(Right(GetPropertyValue("Sprinklers.Sprinkler Start Time"),11))
End Sub

Sub TimeUp (Zone)
	Dim TimeVar
	TimeVar = CInt(GetpropertyValue("Sprinklers.Sprinklers Zone " & Zone & " Time"))
	If TimeVar < 20 Then
		SetpropertyValue "Sprinklers.Sprinklers Zone " & Zone & " Time",  Right("0" & CStr(TimeVar + 1),2)
	End if
End Sub

Sub TimeDown (Zone)
	Dim TimeVar
	TimeVar = CInt(GetpropertyValue("Sprinklers.Sprinklers Zone " & Zone & " Time"))
	If TimeVar > 0 Then
		SetpropertyValue "Sprinklers.Sprinklers Zone " & Zone & " Time",  Right("0" & CStr(TimeVar - 1),2)
	End If	
End Sub


Sub StartTimeUp (TimeInterval)
	Dim TimeVar, IntMax
	If TimeInterval = "Hour" Then
		IntMax = 24
	Elseif 	TimeInterval = "Minute" Then
		IntMax = 60
	End if
	TimeVar = CInt(GetpropertyValue("Sprinklers.Sprinklers Start " & TimeInterval))
	If TimeVar < IntMax Then
		SetpropertyValue "Sprinklers.Sprinklers Start " & TimeInterval, Right("0" & CStr(TimeVar + 1),2)
	End if
End Sub

Sub StartTimeDown (TimeInterval)
	Dim TimeVar
	TimeVar = CInt(GetpropertyValue("Sprinklers.Sprinklers Start " & TimeInterval))
	If TimeVar > 0 Then
		SetpropertyValue "Sprinklers.Sprinklers Start " & TimeInterval, Right("0" & CStr(TimeVar - 1),2)
	End If	
End Sub

Function CountSprinklerZones
	Dim SprinklerZoneCount, NoMoreSprinklerZones
	NoMoreSprinklerZones = 0
	SprinklerZoneCount = 0
	'  Warning -- assumes Sprinkler IDs are sequential and there are no ID gaps
	Do Until NoMoreSprinklerZones = 1
		If GetPropertyValue ("Sprinklers.Sprinklers Zone " & CStr(SprinklerZoneCount+1) + " State") <> "* error *" Then
		   SprinklerZoneCount=SprinklerZoneCount + 1
		Else
			NoMoreSprinklerZones = 1
		End if
	Loop
	CountSprinklerZones = SprinklerZoneCount
End Function




