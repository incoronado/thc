Option Explicit
'On Error Resume Next
Dim socket
Set socket = CreateObject("SimpleSocket.Udp", "event_")
Sub event_OnReceive(data)
	SetpropertyValue "System.Action", data

End Sub

socket.Listen "127.0.0.1", 10000
While(True)
    'WScript.Sleep(5000)
Wend




socket.Abort