<%@ Language=VBScript %>
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	
	On Error Resume Next
	
	strConn = Session("pro_con")
	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "car_shop_request_search_soap_log_select"
	adoCmd.CommandType = 4
	
	Set adoRS = adoCmd.Execute
	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting log information</b><br>"
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
<title>Current Status</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="font-family: Tahoma" >

<p>&nbsp;</p>
<p align="center">Most recent request by machine (PST)</p>
<div align="center">

<table border="0" width="300" id="table1">
	<tr>
		<td bgcolor="#000000"><font color="#FFFFFF">Machine</font></td>
		<td bgcolor="#000000">
		<p align="center"><font color="#FFFFFF">Last Request</font></td>
	</tr>
	
	<% While (adoRS.EOF = False) %>

	<tr>
		<td><font size="2"><%=adoRS.Fields("search_machine").Value %></font></td>
		<td align="center"><font size="2"><%=FormatDateTime(adoRS.Fields("requested").Value, 2) & " " & FormatDateTime(adoRS.Fields("requested").Value, 3) %></font></td>
	</tr>
	
	<% adoRS.MoveNext %>
	<% Wend %>


</table>

<p>&nbsp;</p>
<p><font size="2">* 12:00:00 AM means this value<br>
has not yet been set by the <br>
particular search machine</font></div>

</body>
</html>
<%	
	Rem Clean up
	Set adoCmd = Nothing 
	Set adoRS = Nothing 
	
%>