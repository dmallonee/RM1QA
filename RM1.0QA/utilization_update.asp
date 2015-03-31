<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  
	Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"
	Response.Buffer = True
	Response.Write "<H1>Please Wait .... </H1>"
	Response.Flush

    Server.ScriptTimeout = 3600

	On Error Resume Next

	strUserId = Request.Cookies("rate-monitor.com")("user_id")

	strConn = Session("pro_con")
	
  	Set adoRS = CreateObject("ADODB.Recordset")

	Rem First, update the schedule header, just in case the user changed the description/name
  	Set adoCmd = CreateObject("ADODB.Command")
		
	adoCmd.ActiveConnection = strConn
	adoCmd.CommandTimeout  = 3600
	'adoCmd.CommandType = 4

	adoCmd.CommandText = "EXEC car_utilization_insert_fox2 2, Null"

	Set adoRS = adoCmd.Execute()

	'adoCmd.CommandText = "EXEC car_rental_transaction_res_summary 2"

	'Set adoRS = adoCmd.Execute()

	'adoCmd.CommandText = "EXEC car_rental_transaction_out_summary 2"

	'Set adoRS = adoCmd.Execute()

	'adoCmd.Close 
	
	Set adoRS = Nothing
	Set adoCmd = Nothing

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	Else
		Set adoCmd = Nothing
		Server.Transfer "system_utilization.asp"
	End If

%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor by Rate-Highway, Inc. | Rate Rule Schedule Management</title>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script type="text/javascript" language="javascript" src="inc/sitewide.js" ></script>
</head>
<body topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_middle.jpg">
    <img alt="" src="images/top_left.jpg" width="423" height="91"></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif"><img src="images/top_right.jpg" width="365" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif"></td>
  </tr>
</table>
<p align="right"><font class="copyright">Copyright (c) 2001-<%=Year(Now)%>, Rate-Highway, Inc. All Rights Reserved.</font></p>
<p>
&nbsp;<p align="center">
<font size="4">Error encountered updating Utilization</font><p align="center">
&nbsp;<form method="POST" action="utilization_update.asp" webbot-action="--WEBBOT-SELF--">
	<!--webbot bot="SaveResults" U-File="_private/form_results.csv" S-Format="TEXT/CSV" S-Label-Fields="TRUE" startspan --><input TYPE="hidden" NAME="VTI-GROUP" VALUE="0"><!--webbot bot="SaveResults" i-checksum="43374" endspan -->
	<p align="center"><input type="button" value=" Close window " name="close" onClick="javascript:window.close();">
</p>
</form>
<p align="center">
&nbsp;<p align="center">
&nbsp;<p align="center">
<!--
Please disregard the debug information below<p align="center">
we are currently working on this section to add improvements<p>
&nbsp;<p></p>

-->		
</body>

</html>