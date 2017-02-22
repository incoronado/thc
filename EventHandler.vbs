' EventHandler
' This script is the main dispatcher application
' ©2016 Mike Larson
' 
' Requirements:
' Each Asychronous Event must be subscibed 
' Subscription Requires a HB Null Device using the following format
' Null Device Name Must be "Subscriber-N" Where N is the Unique Subscriber Index
' Subscriber Indexes must be sequemtial and can have no gaps
' Usage:
' 
' Handle the events.
' Message Format Class.Source.Priority.Command:Param1:Param2
'System.GalaxyProTab1.10.TestCommand
' 
'
' Limitations:
' 
' 
Option Explicit
'On Error Resume Next

Dim MessageHandlerVar, EventHandlerCount, DequeueEvent, Queue, i, SleepVar
Dim HandlerQueue1, HandlerQueue2, HandlerQueue3, HandlerQueue4, HandlerQueue5, HandlerQueue6
Dim HandlerQueue7, HandlerQueue8, HandlerQueue9, HandlerQueue10, HandlerQueue11, HandlerQueue12



EventHandlerCount = CountEventHandlers
'Initialize Event Subscriber Queues
'ClearEventHandlerData EventHandlerCount, Queue

Set HandlerQueue1 = CreateObject("System.Collections.Queue")
Set HandlerQueue2 = CreateObject("System.Collections.Queue")
Set HandlerQueue3 = CreateObject("System.Collections.Queue")
Set HandlerQueue4 = CreateObject("System.Collections.Queue")
Set HandlerQueue5 = CreateObject("System.Collections.Queue")
Set HandlerQueue6 = CreateObject("System.Collections.Queue")
Set HandlerQueue7 = CreateObject("System.Collections.Queue")
Set HandlerQueue8 = CreateObject("System.Collections.Queue")
Set HandlerQueue9 = CreateObject("System.Collections.Queue")
Set HandlerQueue10 = CreateObject("System.Collections.Queue")
Set HandlerQueue11 = CreateObject("System.Collections.Queue")
Set HandlerQueue12 = CreateObject("System.Collections.Queue")


ClearEventQueues(EventHandlerCount)

SleepVar = CInt(GetPropertyValue("System.Handler Sleep Time"))
'SetPropertyValue "System.Debug", EventHandlerCount

Do
	For i = 1 To EventHandlerCount
		' Is there a New Message
		MessageHandlerVar = GetPropertyValue ("Handler-" & CStr(i) & ".HandlerMessage")
		If MessageHandlerVar <> "Idle" Then
			SetPropertyValue "Handler-" & CStr(i) & ".HandlerMessage", "Idle"
			EnqueueMessage i, MessageHandlerVar
		End If
		'Process 
		
		If CInt(QueueCount(i)) > 0 Then
			If GetPropertyValue(GetPropertyValue ("Handler-" & CStr(i) + ".MessageHandler")) = "Idle" Then
				SetPropertyValue GetPropertyValue ("Handler-" & CStr(i) + ".MessageHandler"), DequeueMessage(i)
			End if
		End if
	Next
	Sleep SleepVar
Loop



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

Function ClearEventQueues(HandlerCount)
' Clears/Initializes all values in event subscriber queues
	Dim j
	For j = 1 To HandlerCount
		SetpropertyValue "Handler-" + CStr(j) + ".FinalizeNotifier", "False"
		SetpropertyValue "Handler-" + CStr(j) + ".CallbackMessage", ""
		SetpropertyValue "Handler-" + CStr(j) + ".HandlerMessage", "Idle"
		SetpropertyValue "Handler-" + CStr(j) + ".QueueCount", 0
		SetpropertyValue "Handler-" + CStr(j) + ".MaxQueueCount", 0
	Next

End Function

Function EnqueueMessage(QueueID, Message)
	Dim QueueLength
	Select Case QueueID
		Case 1
			HandlerQueue1.Enqueue(Message)
			QueueLength=HandlerQueue1.Count
		Case 2
			HandlerQueue2.Enqueue(Message)
			QueueLength=HandlerQueue2.Count
		Case 3
			HandlerQueue3.Enqueue(Message)
			QueueLength=HandlerQueue3.Count
		Case 4
			HandlerQueue4.Enqueue(Message)
			QueueLength=HandlerQueue4.Count
		Case 5
			HandlerQueue5.Enqueue(Message)
			QueueLength=HandlerQueue5.Count
		Case 6
			HandlerQueue6.Enqueue(Message)
			QueueLength=HandlerQueue6.Count
		Case 7
			HandlerQueue7.Enqueue(Message)
			QueueLength=HandlerQueue7.Count
		Case 8
			HandlerQueue8.Enqueue(Message)
			QueueLength=HandlerQueue8.Count
		Case 9
			HandlerQueue9.Enqueue(Message)
			QueueLength=HandlerQueue9.Count
		Case 10
			HandlerQueue10.Enqueue(Message)
			QueueLength=HandlerQueue10.Count
		Case 11
			HandlerQueue11.Enqueue(Message)
			QueueLength=HandlerQueue11.Count
		Case 12
			HandlerQueue12.Enqueue(Message)
			QueueLength=HandlerQueue12.Count			
	End Select	
	SetpropertyValue "Handler-" + CStr(QueueID) + ".QueueCount", QueueLength
	'Record Max Queue Length
	If QueueLength >  CInt(GetpropertyValue("Handler-" + CStr(QueueID) + ".MaxQueueCount")) Then
	  SetpropertyValue "Handler-" + CStr(QueueID) + ".MaxQueueCount", QueueLength
	End If
	
End Function

Function DequeueMessage(QueueID)
    Dim ReturnStr, QueueLength
	ReturnStr = ""
	Select Case QueueID
		Case 1
			ReturnStr=HandlerQueue1.Dequeue()
			QueueLength=HandlerQueue1.Count
		Case 2
			ReturnStr=HandlerQueue2.Dequeue()
			QueueLength=HandlerQueue2.Count
		Case 3
			ReturnStr=HandlerQueue3.Dequeue()
			QueueLength=HandlerQueue3.Count
		Case 4
			ReturnStr=HandlerQueue4.Dequeue()
			QueueLength=HandlerQueue4.Count
		Case 5
			ReturnStr=HandlerQueue5.Dequeue()
			QueueLength=HandlerQueue5.Count
		Case 6
			ReturnStr=HandlerQueue6.Dequeue()
			QueueLength=HandlerQueue6.Count
		Case 7
			ReturnStr=HandlerQueue7.Dequeue()
			QueueLength=HandlerQueue7.Count
		Case 8
			ReturnStr=HandlerQueue8.Dequeue()
			QueueLength=HandlerQueue8.Count
		Case 9
			ReturnStr=HandlerQueue9.Dequeue()
			QueueLength=HandlerQueue9.Count
		Case 10
			ReturnStr=HandlerQueue10.Dequeue()
			QueueLength=HandlerQueue10.Count
		Case 11
			ReturnStr=HandlerQueue11.Dequeue()
			QueueLength=HandlerQueue11.Count
		Case 12
			ReturnStr=HandlerQueue12.Dequeue()
			QueueLength=HandlerQueue12.Count			
	End Select	
	SetpropertyValue "Handler-" + CStr(QueueID) + ".QueueCount", QueueLength
	DequeueMessage = ReturnStr
	
End Function

Function QueueCount(QueueID)
    Dim QueueLength
	
	Select Case QueueID
		Case 1
			QueueLength=HandlerQueue1.Count
		Case 2
			QueueLength=HandlerQueue2.Count
		Case 3
			QueueLength=HandlerQueue3.Count
		Case 4
			QueueLength=HandlerQueue4.Count
		Case 5
			QueueLength=HandlerQueue5.Count
		Case 6
			QueueLength=HandlerQueue6.Count
		Case 7
			QueueLength=HandlerQueue7.Count
		Case 8
			QueueLength=HandlerQueue8.Count
		Case 9
			QueueLength=HandlerQueue9.Count
		Case 10
			QueueLength=HandlerQueue10.Count
		Case 11
			QueueLength=HandlerQueue11.Count
		Case 12
			QueueLength=HandlerQueue12.Count			
	End Select	
	SetpropertyValue "Handler-" + CStr(QueueID) + ".QueueCount", QueueLength
	QueueCount = QueueLength
End Function



