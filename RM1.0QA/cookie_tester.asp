<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 360

	On Error Resume Next
 


%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor by Rate-Highway, Inc. | Rate-Change Report</title>
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
<p align="right"><font class="copyright">Copyright (c) 2001-<%=Year(Now)%>, Rate-Highway, Inc. All Rights Reserved.</font></p>
<p>
<p align="center"><font size="5" color="#384F5B">Your cookies have been 
collected.</font></p>
<p align="center">&nbsp;</p>
<form method="POST" action="" name="refresh">
  <p align="center"><input name="refresh" type="submit" value="Refresh"></p>
  
</form>
<p><font size="-1"><br>
&nbsp;</font></p>

<%

	'Response.Cookies("rate-monitor.com").Domain =  Request.ServerVariables("SERVER_NAME")	 		
	'			Response.Cookies("rate-monitor.com").Path = "/"	 		
	'	 		Response.Cookies("rate-monitor.com")("live_session") = "auto"
	'			Response.Cookies("rate-monitor.com").Expires = Now + 1

			Response.write "Begin<br>"	

			Response.write Request.Cookies("rate-monitor.com").Domain & "<br>"	
			Response.write Request.Cookies("rate-monitor.com").Path & "<br>"	 		

			Response.write Session("pro_con") & "<br>"
	 		Response.write Request.Cookies("rate-monitor.com")("live_session") & "<br>"
	 		Response.write Request.Cookies("rate-monitor.com")("loginCode") & "<br>" 
			Response.write Request.Cookies("rate-monitor.com")("password") & "<br>"
			Response.write Request.Cookies("rate-monitor.com")("remember Me") & "<br>"
			Response.write Request.Cookies("rate-monitor.com")("testing") & "<br>"
			Response.write Request.Cookies("rate-monitor.com")("user_id") & "<br>"
			Response.write Request.Cookies("rate-monitor.com")("user_name") & "<br>"
			Response.write Request.Cookies("rate-monitor.com")("client_userid") & "<br>"
			Response.write Request.Cookies("rate-monitor.com").Expires & "<br>"
			Response.write Request.Cookies("rate-monitor.com")("vend_cd") & "<br>"	
			Response.write Request.Cookies("rate-monitor.com")("rpt_limit") & "<br>"
			Response.write Now + 1  & "<br>End<br>"	

%>
</body>

</html>