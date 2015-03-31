<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 360

	'On Error Resume Next
 
	Response.Cookies("rate-monitor.com") = ""

%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor by Rate-Highway, Inc. | Rate Rule Upload</title>
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
.style1 {
	font-size: large;
	color: #384F5B;
}
.style2 {
	text-align: center;
}
.style3 {
	background-color: #E1DFCC;
}
.style4 {
	font-size: x-small;
}
.style5 {
	font-size: x-small;
	text-align: left;
}
-->
</style>
</head>
<body topmargin="0">
<div class="style2">
<div class="style2">
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
<p align="center" class="style1">Rate R<font size="5">ule Upload</font></p>
	<div class="style2">
		
			
			<span class="style3">
			<OBJECT WIDTH=500 HEIGHT=200 
  ID="UploadCtl" 
  CLASSID="CLSID:E87F6C8E-16C0-11D3-BEF7-009027438003" 
  CODEBASE="rule_upload/XUpload.ocx#VERSION=3,0,0,0">
				<PARAM NAME="RegKey" value="zBrxP7IrU54rhoEIzZduDBu564vunqlS9SSVoZNo900KlM3DDPjwT6HMLcX9QPCKPgtFKz3WNZp2">
				<param name="server" value="www.rate-monitor.com">
				<param name="script" value="/rule_upload/01_simple_upload.asp">
				<PARAM NAME="Filter" VALUE="Excel CSV Files (.csv)|*.csv">
			</OBJECT>
<!-- Microsoft workaround for the Click-to-Activate problem -->
<script type="text/javascript" src="rule_upload/ie_workaround.js"></script>

			</span><br>
			<br>
			<table style="width: 493px" align="center">
				<tr>
					<td style="width: 29px" class="style5" valign="top">1)</td>
					<td class="style4">Right click and select the file you would like to upload. (repeat for as many files as you would like to upload).</td>
				</tr>
				<tr>
					<td style="width: 29px" class="style5" valign="top">2)</td>
					<td class="style4">Once all files have been selected, Right-click again and select 
			upload&nbsp;.</td>
				</tr>
				<tr>
					<td style="width: 29px" class="style4">3)&nbsp;</td>
					<td class="style4">Once the files are uploaded, you may close this window.</td>
				</tr>
			</table>
			<br>
			<p align="center">
			<input type="button" value=" Close " name="close"  onClick="javascript:window.close();" class="rh_button"></p>


			</div>
</div>
</div>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>