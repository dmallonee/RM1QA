<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 360

%>

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor by Rate-Highway, Inc. | Support files you can download </title>
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

<SCRIPT LANGUAGE="JavaScript"  >
<!-- Begin
function checkAll(field)
{
	alert(field.name);

	for (i = 0; i < field.length; i++)
		field[i].checked = true ;
}

function uncheckAll(field)
{

	for (i = 0; i < field.length; i++)
		field[i].checked = false ;
}
//  End -->
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
	font-size: large;
	color: #384F5B;
}
.style2 {
	font-size: x-small;
}
.style3 {
	border-width: 0px;
}
.style4 {
	text-align: center;
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
<font size="-1">
 <table width="1000" border="0" cellpadding="2" style="border-collapse: collapse" id="table0">
    <tr valign="bottom">
      <td >
      <p align="center" class="style1">S<font size="5">upport Documents</font></td>
    </tr>
    <tr valign="bottom">
      <td >&nbsp;</td>
    </tr>
  </table>
  </font>

    
 
	  	
		<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1000" height="4" id="footerbar<%=intTableCount %>">
		    <tr>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		      <td background="images/ruler.gif"></td>
		    </tr>
		</table>
		<br>
		
<br>
<table style="width: 700px" align="center">
	<tr class="profile_header">
		<td style="width: 130px">Name</td>
		<td style="width: 390px">Description</td>
		<td>(right-click to download)</td>
	</tr>
	<tr>
		<td style="width: 130px" class="style2" valign="top"><strong>User Guide</strong><br>
		(PDF format)</td>
		<td class="style2" style="width: 390px" valign="top">Rate-Monitor user 
		manual. Each section of the system is covered in this document intended 
		for new users wanting a better overall understanding and for existing 
		users that need a reference.</td>
		<td class="style2" valign="top">
		<a title="right click and select save as" href="docs/Rate-Monitor%20User%20Guide.pdf">
		user guide</a> (1.8 MBytes)</td>
	</tr>
	<tr>
		<td style="width: 130px" class="style2" valign="top"><strong>Schedule 
		min-max</strong><br>
		(Excel XLS format)</td>
		<td class="style2" style="width: 390px" valign="top">Rule schedule 
		min-max setup form. Download and fill out your min-max numbers then send 
		to Support if you want this configured for you.</td>
		<td class="style2" valign="top">
		<a title="right click and select save as to download this document" href="docs/Schedule%20Min-Max%20setup.xls">
		Min/Max</a> (181 KBytes)</td>
	</tr>
	<tr>
		<td style="width: 130px" class="style2">&nbsp;</td>
		<td class="style2" style="width: 390px" valign="top">&nbsp;</td>
		<td class="style2">&nbsp;</td>
	</tr>
</table>
<p>&nbsp;</p>
<p class="style4">&nbsp;<a target="_blank" href="http://www.adobe.com/products/acrobat/readstep2.html"><img alt="Download Adobe Acrobat Reader" longdesc="Click here to download the Adobe Acrobat Reader" src="images/get_adobe_reader.gif" width="112" height="33" class="style3"></a>
</p>
<p class="style4">Need to view PDF files and you don't have Acrobat? Click the 
button to download.</p>
<p class="style4">&nbsp;</p>
<!--#INCLUDE FILE="footer.asp"-->		
</body>
</html>