<%@ Language=VBScript %>
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

	%>

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor by Rate-Highway, Inc. | Rate Changes Complete</title>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="javascript" src="inc/sitewide.js" ></script>
<script language="javascript" src="inc/header_menu_support.js"></script>
<script language='Javascript'> 
	function centerPopUp( url, name, width, height, scrollbars ) { 
 
	if( scrollbars == null ) scrollbars = "0" 
 
	str  = ""; 
	str += "resizable=1,"; 
	str += "scrollbars=" + scrollbars + ","; 
	str += "width=" + width + ","; 
	str += "height=" + height + ","; 
    
	if ( window.screen ) { 
		var ah = screen.availHeight - 30; 
		var aw = screen.availWidth - 10; 
 
		var xc = ( aw - width ) / 2; 
		var yc = ( ah - height ) / 2; 
 
		str += ",left=" + xc + ",screenX=" + xc; 
		str += ",top=" + yc + ",screenY=" + yc; 
	} 
	window.open( url, name, str ); 
} 
</script> 

<style>
<!--
P {
	COLOR: navy; FONT-FAMILY: Verdana, Arial, sans-serif
}
.copyright {
	FONT-SIZE: 0.7em; TEXT-ALIGN: right
}
.style1 {
	text-align: center;
	font-family: tahoma,verdana,arial,helvetica,sans-serif;
	font-size: x-small;
	color: black;
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
<p align="center"><font size="5" color="#384F5B">Thank you.</font></p>
<p align="center"><font size="5" color="#384F5B">Your rate changes have been 
accepted.<br>
&nbsp;</font></p>
<form method="POST" action="return false" name="close">
  <p align="center">
  <input type="button" value=" Close " name="close"  onClick="javascript:window.close();" class="rh_button">
</p></form>
<form action="rate_change_report.asp" method="post" name="more_changes" >
	<p align="center">&nbsp;</p>
	<p align="center"><span class="style1">Or, if you need to change more rates from the same report, 
	click the more changes button below:<br><br><strong>Please Note</strong>: It 
	may take up to five minutes for the changes you just submitted to be<br>
	processed and appear on the suggestion report with an updated status.<br>
	</span><br>
	<input name="more" type="submit" value="More Changes" class="rh_button"></p>
	<input name="reportrequestid" type="hidden" value='<%=Session("reportrequestid") %>'>
	<input name="security_code" type="hidden" value='<%=Session("security_code") %>'>
</form>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>