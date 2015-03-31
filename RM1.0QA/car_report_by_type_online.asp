<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 360

	'On Error Resume Next
 
	Dim strConn	
	Dim adoCmd	
	Dim adoRS


	If IsNumeric(Request("reportrequestid")) Then

		strConn = Session("pro_con")
	
	  	Set adoRS = CreateObject("ADODB.Recordset")
	  	Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "validate_security_code"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request("reportrequestid"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@guid",           72, 1, 0, Request("security_code"))

		Set adoRS = adoCmd.Execute

	Else
		Server.Transfer "security_code_failed.asp"	
	
	End If
	
	
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

		If (adoRS.EOF = False) Then
		
			If IsNull(adoRS.Fields("shop_request_id").Value) = False Then
				Response.Cookies("rate-monitor.com").Domain =  Request.ServerVariables("SERVER_NAME")	 		
				Response.Cookies("rate-monitor.com").Path = "/"	 		
				Response.Cookies("rate-monitor.com").Expires = DateAdd("h", 1, Now)
				Response.Cookies("rate-monitor.com")("live_session") = "auto"


			Else
				Server.Transfer "security_code_failed.asp"
			
			End If
	
		Else
			Server.Transfer "security_code_failed.asp"
		
		End If	
	
	
	End If


%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor.com | Rate Report</title>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="javascript" src="inc/sitewide.js" ></script>
<script language="javascript" src="inc/header_menu_support.js"></script>
<style>
<!--
P {
	COLOR: navy; FONT-FAMILY: Verdana, Arial, sans-serif
}
.copyright {
	FONT-SIZE: 0.7em; TEXT-ALIGN: right
}
-->
</style>
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
<p align="center"><font size="5" color="#384F5B">Your security code has been 
confirmed.</font></p>
<p align="center">&nbsp;</p>
<form method="POST" action="car_report_by_type.asp" name="rate_report">
  <p align="center"><input type="submit" value="View Report" name="view_report"></p>
  <input type="hidden" name="reportrequestid" value="<%=Request("reportrequestid") %>">
  <input type="hidden" name="security_code" value="<%=Request("security_code") %>">
</form>
<p><font size="-1"><br>
&nbsp;</font></p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>