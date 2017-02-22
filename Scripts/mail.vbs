Option Explicit
On Error Resume Next

Dim Action
Dim SleepVar

SleepVar = CInt(GetPropertyValue("System.Script Sleep Time"))

Do
	Sleep SleepVar
		
	Action = GetPropertyValue ("Mailer.Action")
	If Action <> "Idle" Then
		SetpropertyValue "Mailer.Action", "Idle"
		Sleep SleepVar
		If Action <> "" Then
			Call ProcessMailMessage(Action)
		End If
	End If
Loop


Sub ProcessMailMessage(Action)
    'MsgBox "Ready Command"
    Select Case Action 
	' Configuration Command
	   Case "SendSMS"
	     SendMail
	   End Select
End Sub


Sub SendMail
	Dim EmailSubject, EmailBody, EmailFrom, EmailFromName, EmailTo, SMTPServer, SMTPLogon, SMTPPassword, objMessage
	
	EmailSubject = GetPropertyValue ("Mailer.MailSubject")
	EmailBody = GetPropertyValue ("Mailer.MailBody")

	EmailFrom = GetPropertyValue ("Mailer.MailFrom")
	EmailFromName = GetPropertyValue ("Mailer.FromName")
	EmailTo = GetPropertyValue ("Mailer.MailRecipients")
	SMTPServer = GetPropertyValue ("Mailer.SMTPServer")
	SMTPLogon = GetPropertyValue ("Mailer.SMTPLogin")
	SMTPPassword = GetPropertyValue ("Mailer.SMTPPassword")
	Const SMTPSSL = True
	Const SMTPPort = 465

	Const cdoSendUsingPickup = 1 	'Send message using local SMTP service pickup directory.
	Const cdoSendUsingPort = 2 	'Send the message using SMTP over TCP/IP networking.

	Const cdoAnonymous = 0 	' No authentication
	Const cdoBasic = 1 	' BASIC clear text authentication
	Const cdoNTLM = 2 	' NTLM, Microsoft proprietary authentication

	' First, create the message

	Set objMessage = CreateObject("CDO.Message")
	objMessage.Subject = EmailSubject
	objMessage.From = """" & EmailFromName & """ <" & EmailFrom & ">"
	objMessage.To = EmailTo
	objMessage.TextBody = EmailBody

	' Second, configure the server

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPLogon

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

	objMessage.Configuration.Fields.Update

	' Now send the message!

	objMessage.Send
End Sub