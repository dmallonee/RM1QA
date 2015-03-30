<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180
   
    On Error Resume Next
 
	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim strResult 

	'Retrieve the name of the current ASP document
	'sPageURL = Request.ServerVariables("SCRIPT_NAME")
	
	strConn = Session("pro_con")
	
  	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "support_status_check"
	adoCmd.CommandType = adCmdStoredProc 

	Set adoRS = adoCmd.Execute

	If adoRS.EOF = True Then
		strResult = "Server unavailable" 

	Else
		adoRS.Close
		strResult = "Server available" 
	
	End If

	Set adoRS = Nothing
	Set	adoCmd = Nothing

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Status</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>
<body>
<p>Current status: <%=strResult %></p>
</body>
</html>
