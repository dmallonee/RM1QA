<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 360

	On Error Resume Next
	
	
	Rem If the control is disabled it will pass back a null, but we need an empty string instead.
	strTSDCustomerNumber = "" & Request.Form("tsd_customer_number")
	strTSDPasscode       = "" & Request.Form("tsd_passcode")

	strUserId = Request.Cookies("rate-monitor.com")("user_id")

	strConn = Session("pro_con")
	
  	Set adoRS = CreateObject("ADODB.Recordset")

	Rem First, update the schedule header, just in case the user changed the description/name
  	Set adoCmd = CreateObject("ADODB.Command")
		
	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "org_update"
	adoCmd.CommandType = 4

    if Request.Form("enable_rule_processing") = "" then
        enable_rule_processing = "0"
    else
        enable_rule_processing = "1"
    end if
	adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id",               3, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",              3, 1, 0, strUserId)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@update_rezcentral",   11, 1, 0, Request.Form("update_rezcentral"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@update_bluebird",     11, 1, 0, Request.Form("update_bluebird"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@update_tsd",          11, 1, 0, Request.Form("update_tsd"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@ftp_client_id",        3, 1, 0, Request.Form("ftp_client_id"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@time_zone_offset",     2, 1, 0, Request.Form("time_zone_offset"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@no_show_percentage",   6, 1, 0, Request.Form("no_show_percentage"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@utilization_input_id", 3, 1, 0, Request.Form("utilization_input_id"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@weekly_lor",           3, 1, 0, Request.Form("weekly_lor"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@tsd_customer_number",  200, 1, 20, strTSDCustomerNumber )
	adoCmd.Parameters.Append adoCmd.CreateParameter("@tsd_passcode",         200, 1, 20, strTSDPasscode )
	adoCmd.Parameters.Append adoCmd.CreateParameter("@enable_rule_processing", 11, 1, 0, enable_rule_processing)
    		
	adoCmd.Execute

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
<script language="javascript" src="inc/sitewide.js" ></script>
</head>
<body topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_middle.jpg">
    <img src="images/top_left.jpg" width="423" height="91"></td>
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
&nbsp;<form method="POST" action="system_update.asp" webbot-action="--WEBBOT-SELF--">
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
&nbsp;<p>

<%=Request.Form("cell0") %>0<br >
<%=Request.Form("cell1") %>1<br >
<%=Request.Form("cell2") %>2<br >
<%=Request.Form("cell3") %>3<br >
<%=Request.Form("cell4") %>4<br >
<%=Request.Form("cell5") %>5<br >
<%=Request.Form("cell6") %>6<br >
<%=Request.Form("cell7") %>7<br >
<%=Request.Form("cell8") %>8<br >
<%=Request.Form("cell9") %>9<br >
<%=Request.Form("cell10") %>10<br >
<%=Request.Form("cell11") %>11<br >
<%=Request.Form("cell12") %>12<br >
<%=Request.Form("cell33") %>33<br >		
Request.Form("days_out_grp3") = <%=Request.Form("days_out_grp3") %><br>
<p>ScheduleID = <%=intScheduleID %></p>
-->		
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>