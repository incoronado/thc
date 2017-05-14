' Menu Class
' This Script Provides Menu Functions to Housebot
Option Explicit
'On Error Resume Next

Dim Action
Dim SleepVar

Dim HousebotLocation, DOSHousebotLocation 

DOSHousebotLocation = "C:\Progra~1\Housebot\"
HousebotLocation = "C:\Program Files (x86)\housebot\"


'-------------------------------------------------------
'- Main: Checks Received Data and Handles Send Data ----
'-------------------------------------------------------
SleepVar = 5

Do
	Action = GetPropertyValue ("Menu.Action")
	If Action <> "Idle" Then
		SetpropertyValue "Menu.Action", "Idle"
		If Action <> "" Then
			Call Message_Handler(Action)
		End If
	End If
	Sleep SleepVar
Loop

Sub Message_Handler(message)

	dim a, b
    ' menuname.function
	a=split(Action,".")
	Select Case a(0)
		' System Class Message Received
		Case "Menu"
			b=split(a(3),":")
			'Menu.GalaxyTabA1.10.SelectZone:1
			Select Case b(0)
				Case "NextItem"
					'SetPropertyValue "Menu.Debug", "NextItem"
					NextItem (a(0))	
				Case "PreviousItem"
					'SetPropertyValue "Menu.Debug", "NextItem"
					PreviousItem (a(0))
				Case "SelectItem"
					SelectItem a(0), a(2)
				Case "PopulateMenu"	
					PopulateMenu(a(0))
				Case "SelectZone"
					SelectZone b(1), GetRemoteNumber(a(1))
			
			End Select
	End Select	
End Sub



Sub SelectZone(ZoneNo, Remote)
	'Zone: 1=Living Room; 5=Master Bedroom; 6=Patio
	'Zone Name
	Dim status_str, sel_str, i
	SetPropertyValue Remote  & ".Selected Zone Name",  GetPropertyValue("Multiroom Audio Settings.Zone " & CStr(ZoneNo)  & " Name")
	SetPropertyValue "Subscriber-1.DispatchMessage", "System." & GetPropertyValue(Remote & ".Remote Name")  & ".10.SelectZone:" & CStr(ZoneNo)
	
	'test
	
    for i = 1 to 6
		If GetPropertyValue("Multiroom Audio Settings.Zone " &  ZoneNo & " Power")  = "On" Then
			status_str = "on"
		Else
			status_str = "off"
		End if
		
		If i = CInt(ZoneNo) Then
			sel_str = "sel"
			SetPropertyValue Remote  & ".Selected Zone Icon",  "Config\Themes\THC 2560x1440\zone" & CStr(i) & "-sel-" & status_str & ".png"
			
		Else
			sel_str = "unsel"	
			
		End If
		SetPropertyValue  Remote & ".Menu Icon " & Cstr(i) , "Config\Themes\THC 2560x1440\zone" & CStr(i) & "-" & sel_str & "-" & status_str & ".png"
	Next
	'sleep 250
	SetPropertyValue "Subscriber-8.DispatchMessage", "System.Desktop.10.ClosePanel:Zone Menu"
	
	
	


End Sub

Function GetRemoteNumber(RemoteName)
	Dim NoMoreRemotes, RemoteCount, RemoteNumber
	RemoteCount = 0
	NoMoreRemotes = 0
	RemoteNumber = ""
	Do Until NoMoreRemotes = 1
		If GetPropertyValue ("Remote-" + CStr(RemoteCount+1) + ".Remote Name") <> "* error *" Then
		   RemoteCount=RemoteCount + 1
			If GetPropertyValue ("Remote-" + CStr(RemoteCount) + ".Remote Name") = RemoteName Then
				RemoteNumber = "Remote-" + CStr(RemoteCount)	
			End if
		Else
			NoMoreRemotes = 1
		End if
	Loop
	GetRemoteNumber = RemoteNumber
End Function



Sub  SelectItem (MenuName, Item)
	Dim MenuPageItem, MenuPageSize, i , MenuSelectedItem, MenuImage, NewMenuImage, SleepAfterWrite
	MenuPageItem = CInt(Item)
	SetPropertyValue "Menu.Debug", Item
	SleepAfterWrite = 5
	
	MenuPageSize = CInt(GetPropertyValue(MenuName + ".Menu Page Size"))
	MenuSelectedItem = CInt(GetPropertyValue(MenuName + ".Menu Selected Item"))
	Sleep SleepAfterWrite
	SetPropertyValue MenuName + ".Menu Page Item", CStr(MenuPageItem)
	
	
	Sleep SleepAfterWrite
	'SetPropertyValue "Menu.Debug", MenuName + ".Menu Page Size"
	
	
	
	For i=1 to 10
	    MenuImage = GetPropertyValue(MenuName + ".Menu Image " + CStr(i))
		If i = MenuPageItem  Then
			If MenuPageItem = MenuSelectedItem Then
			    NewMenuImage = GetPropertyValue(MenuName + ".Menu Active Image")
				if MenuImage <>  NewMenuImage Then
					SetPropertyValue MenuName + ".Menu Image " + CStr(i) , NewMenuImage
					Sleep SleepAfterWrite
				End if	
			Else
				NewMenuImage = GetPropertyValue(MenuName + ".Menu Cursor Image")
			    if MenuImage <>  NewMenuImage Then
					SetPropertyValue MenuName + ".Menu Image " + CStr(i) , NewMenuImage
					Sleep SleepAfterWrite
				End if	
			End if	
		Else
		    If i = MenuSelectedItem Then
				NewMenuImage = GetPropertyValue(MenuName + ".Menu Select Item")
				if MenuImage <>  NewMenuImage Then
					SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Select Item")
					Sleep SleepAfterWrite
				End if		
			
			Else
				NewMenuImage = GetPropertyValue(MenuName + ".Menu Unselected Image")
				if MenuImage <>  NewMenuImage Then
					SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Unselected Image")
					Sleep SleepAfterWrite
				End if	
			End if
		End if
		
		'if (i = MenuSelectedItem) And (MenuPageItem <> MenuSelectedItem) Then
				'NewMenuImage = GetPropertyValue(MenuName + ".Menu Select Item")
				'if MenuImage <>  NewMenuImage Then
						'SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Select Item")
						'Sleep SleepAfterWrite
				'End if		
		'End if
			
	Next

End Sub




Sub  NextItem (MenuName)
	Dim MenuPageItem, MenuPageSize, i , MenuSelectedItem
	'SetPropertyValue "Menu.Debug", MenuName + ".Menu Page Size"
	MenuPageItem = CInt(GetPropertyValue(MenuName + ".Menu Page Item"))
	MenuPageSize = CInt(GetPropertyValue(MenuName + ".Menu Page Size"))
	MenuSelectedItem = CInt(GetPropertyValue(MenuName + ".Menu Selected Item"))

	If  MenuPageItem < MenuPageSize   Then
		MenuPageItem = MenuPageItem + 1
	Else
		MenuPageItem = 1
	End If	
	SetPropertyValue MenuName + ".Menu Page Item", MenuPageItem
	Sleep 5
	
	For i=1 to 10
		If i = MenuPageItem  Then
			If MenuPageItem = MenuSelectedItem Then
				SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Active Image")
			Else
				SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Cursor Image")
			End if	
		Else 
			SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Unselected Image")
		End if
		
		if (i = MenuSelectedItem) And (MenuPageItem <> MenuSelectedItem) Then
				SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Select Item")
		End if
		Sleep 5	
	Next

End Sub

Sub  PreviousItem (MenuName)
	Dim MenuPageItem, MenuPageSize, i,  MenuSelectedItem
	'SetPropertyValue "Menu.Debug", MenuName + ".Menu Page Size"
	MenuPageItem = Cint(GetPropertyValue(MenuName + ".Menu Page Item"))
	MenuPageSize = CInt(GetPropertyValue(MenuName + ".Menu Page Size"))
    MenuSelectedItem = CInt(GetPropertyValue(MenuName + ".Menu Selected Item"))
	
	If  MenuPageItem > 1   Then
		MenuPageItem = MenuPageItem - 1
	Else
		MenuPageItem = MenuPageSize
	End If	
	SetPropertyValue MenuName + ".Menu Page Item", MenuPageItem

	
	Sleep 5
	
	For i=1 to 10
		If i = MenuPageItem  Then
		    If MenuPageItem = MenuSelectedItem Then
				SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Active Image")
			Else
				SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Cursor Image")
			End if	
		Else 
			SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Unselected Image")
		End if
		
		if (i = MenuSelectedItem) And (MenuPageItem <> MenuSelectedItem) Then
				SetPropertyValue MenuName + ".Menu Image " + CStr(i) , GetPropertyValue(MenuName + ".Menu Select Item")
		End if
		
		

		Sleep 5
	Next

End Sub

Sub PopulateMenu(MenuName)
  Dim objDB, su, SQLStr, Row, r, Label
  Set objDB  = CreateObject("newObjects.sqlite3.dbutf8")
  objDB.Open(HousebotLocation & "config\scripts\songs.mlf")
  objDB.BusyTimeout=1000
  objDB.AutoType = True
  objDB.TypeInfoLevel = 4
  Set su = CreateObject("newObjects.utilctls.StringUtilities")
  SQLStr = "select * from menus where menu_name  = '" & MenuName & "'"
  Set r = objDB.Execute(SQLStr)
  For Row = 1 To r.Count 
	'If Len(r(Row)("menu_label"))>15 Then
	'   Label=mid(r(Row)("menu_label"),1,15) + chr(133)
	'Else
	   Label=r(Row)("menu_label")
	'End If   
	SetPropertyValue MenuName + ".Menu Text " + CStr(Row), Label
  Next
  Set r = Nothing
  sleep 500
End Sub
