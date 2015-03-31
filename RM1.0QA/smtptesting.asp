<%@ Language=VBScript %>


<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

	On Error Resume Next

Const cdoSendUsingPickup = 1
Const cdoSendUsingPort = 2
Const cdoSendUsingExchange = 3

Const cdoAnonymous = 0
Const cdoBasic = 1
Const cdoNTLM = 2

'Sends an email To aTo email address, with Subject And TextBody.
'The email is In text format.
'Lets you specify BCC adresses, Attachments, smtp server And Sender email address
Function SendMailByCDO(aTo, Subject, TextBody, BCC, Files, smtp, aFrom )
  on error resume Next

  Dim Message 'As New CDO.Message '(New - For VBA)
  
  'Create CDO message object
  Set Message = CreateObject("CDO.Message")

  'Set configuration fields. 
  With Message.Configuration.Fields
    'Original sender email address 
    .Item("http://schemas.microsoft.com/cdo/configuration/sendemailaddress") = aFrom

    'SMTP settings - without authentication, using standard port 25 on host smtp
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtp

    'SMTP Authentication
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoAnonymous

    .Update
  End With

  'Set other message fields.
  With Message
    'From, To, Subject And Body are required.
    .From = aFrom
    .To = aTo
    .Subject = Subject

    'Set TextBody property If you want To send the email As plain text
    .TextBody = TextBody

    'Set HTMLBody  property If you want To send the email As an HTML formatted
    '.HTMLBody = TextBody

    'Blind copy And attachments are optional.
    If Len(BCC)>0 Then .BCC = BCC
    If Len(Files)>0 Then .AddAttachment Files
    
    'Send the email
    .Send
  End With

  'Returns zero If succesfull. Error code otherwise 
  SendMailByCDO = Err.Number
End Function

'Send one email with two BCC addresses And file attachment
'using smtp.mycompany.To
SendMail "michaelm@rate-highway.com", "Subject of the message", "Some interesting plain text body", "", ""

Function SendMail(ByVal aTo, ByVal Subject, ByVal TextBody, ByVal BCC, byref Files )
  Const smtp = "zeus.rate-monitor.com"
  Const aFrom = "SMTP Testing <smtp@rate-highway.com>"
  SendMail = SendMailByCDO(aTo, Subject, TextBody, BCC, Files, smtp, aFrom )
End Function


        If err.number <> 0 Then
	        pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	        response.write "<b>VBScript Errors Occured!<br>"
	        response.write "</b><br>"
	        response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	        response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	        response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	        response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	        response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
        End If



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" />
<title>Untitled 1</title>
</head>

<body>


			   <%	For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"

       
					Next
				%>

</body>

</html>
