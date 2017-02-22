' Dispatcher
' This script is the main dispatcher application
' ©2016 Mike Larson
' 
' Requirements:
' Requires 
'
' Usage:
' 
' 
' Handle the Dispatchs.
' Destroy to release system resources when no longer required.
'
' Limitations:
' Does not support multi-processor systems.

Option Explicit
'On Error Resume Next

Dim IncomingMessage, DispatchSubscriberCount, HandlerCount, i, SleepVar

HandlerCount = CountEventHandlers
DispatchSubscriberCount = CountDispatchSubscribers

'MsgBox HandlerCount

SleepVar = CInt(GetPropertyValue("System.Dispatch Sleep Time"))
Do
	For i = 1 To DispatchSubscriberCount
		IncomingMessage = Trim(GetPropertyValue ("Subscriber-" + CStr(i) + ".DispatchMessage"))
		If IncomingMessage <> "Idle" Then
			SetpropertyValue "Subscriber-" + CStr(i) + ".DispatchMessage", "Idle"
			BlockSubscriber(i)	
			DispatchMessage IncomingMessage, HandlerCount
			'SetpropertyValue "DispatcherScript.Debug", "Idle"
			sleep 5
			UnBlockSubscriber(i)
		End If
	Next
	Sleep SleepVar
Loop

Sub DispatchMessage(MessageStr, HandlerCount)
	Dim a, i
	'Message Format Class.Source.Priority.Command:Param1:Param2
	a=split(MessageStr,".")
	'SetpropertyValue "DispatcherScript.Debug", a(0)
	For i = 1 To HandlerCount
	    'MsgBox a(0) & "=" & GetPropertyValue ("Handler-" + CStr(i) + ".Class")
		If a(0) = GetPropertyValue ("Handler-" + CStr(i) + ".Class") Then
			SetpropertyValue "Handler-" + CStr(i) + ".HandlerMessage", MessageStr
		End if
	Next
End Sub		

Function CountDispatchSubscribers
	Dim DispatchSubscriberCount,NoMoreDispatchSubscribers
	NoMoreDispatchSubscribers = 0
	DispatchSubscriberCount = 0
	'  Warning -- assumes light IDs are sequential and there are no ID gaps
	Do Until NoMoreDispatchSubscribers = 1
		If GetPropertyValue ("Subscriber-" + CStr(DispatchSubscriberCount+1) + ".Name") <> "* error *" Then
			DispatchSubscriberCount=DispatchSubscriberCount + 1
		Else
			NoMoreDispatchSubscribers = 1
		End if
	Loop
	CountDispatchSubscribers = DispatchSubscriberCount
End Function

Function CountEventHandlers
	Dim EventHandlerCount, NoMoreEventHandlers
	NoMoreEventHandlers = 0
	EventHandlerCount = 0
	'  Warning -- assumes light IDs are sequential and there are no ID gaps
	Do Until NoMoreEventHandlers = 1
		If GetPropertyValue ("Handler-" + CStr(EventHandlerCount+1) + ".Class") <> "* error *" Then
		   EventHandlerCount=EventHandlerCount + 1
		Else
			NoMoreEventHandlers = 1
		End if
	Loop
	CountEventHandlers = EventHandlerCount
End Function

Sub BlockSubscriber(SubscriberID) 
	SetModeState "Subscriber-" & CStr(SubscriberID), "Inactive"
End Sub	

Sub UnBlockSubscriber(SubscriberID) 
	SetModeState "Subscriber-" & CStr(SubscriberID), "Active"
End Sub	

		