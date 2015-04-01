<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 60

	'On Error Resume Next

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoCmd1	
	Dim adoRS1
	Dim adoCmd2	
	Dim adoRS2
	Dim adoCmd3	
	Dim adoRS3
	Dim adoCmd4
	Dim adoRS4


	Dim adoPrices
	Dim strUserId
	Dim intRuleId
	Dim strAlertDesc
	Dim datBeginDate
	Dim intComparisonRate 
	Dim strRateAmtTolerance
	Dim strRuleStatus
	
	'Declare variables
	Dim iCurrentPage
	Dim intPageSize
	Dim i
	Dim oConnection
	Dim oRecordSet
	Dim oTableField
	Dim sPageURL


	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	Rem Get the schedules

	Set adoCmd9 = CreateObject("ADODB.Command")

	adoCmd9.ActiveConnection =  strConn
	adoCmd9.CommandText = "car_rate_rule_schedule_select"
	adoCmd9.CommandType = adCmdStoredProc

	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@org_id",              3, 1, 0, Null)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@schedule_type_id",    3, 1, 0, 4)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@user_id",             3, 1, 0, strUserId)
		
	Set adoRS10 = adoCmd9.Execute
	
	Set adoCmd9 = Nothing	
	
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

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor by Rate-Highway, Inc. | Max / Min Schedule</title>
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
.data_cell   { width: 65; text-align: right; font-family: Tahoma; font-size: 10pt }
.header      { width: 65; text-align: center; background-color: #CFD7DB }
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
<p align="center"><font size="5" color="#384F5B">Rate Max. / Min. Schedule</font></p>
<p align="center"><font size="2" color="#384F5B">Edit an existing schedule by 
selecting one from the list below</font></p>
<form method="POST" action="rate_rule_maxmin_schedule_b.asp?type=existing" name="rate_rule_maxmin">
 
	<div align="center">
 
	<table border="0" width="640" id="table1" bgcolor="#CFD7DB">
		<tr>
			<td width="128"><font size="2">Edit Schedule: </font></td>
			<td>&nbsp;</td>
			<td width="493">
			<select type="text" name="car_rate_rule_schedule_id" style="width: 400" size="1"  >
     		<% While adoRS10.EOF = False %>
		   		<option value="<%=adoRS10.Fields("car_rate_rule_schedule_id").Value %>"><%=adoRS10.Fields("car_rate_rule_schedule_desc").Value  %></option>
	   			<% adoRS10.MoveNext %>
     		<% Wend             %>
			</select>
			</td>
		</tr>
		<tr>
			<td width="128">&nbsp;</td>
			<td>&nbsp;</td>
			<td width="493">&nbsp;</td>
		</tr>
		<tr>
			<td width="128">&nbsp;</td>
			<td>&nbsp;</td>
			<td width="493"> <input type="submit" value="  Next &gt;&gt;  " name="B1"></td>
		</tr>
	</table>
	</div>
  <p align="center">&nbsp;</p>
</form>
<p align="center"><font color="#384F5B">Or create a new schedule by using the 
option below</font></p>
<form method="POST" action="rate_rule_maxmin_schedule_b.asp?type=new" name="new_rate_rule_maxmin_schedule">
 
	<div align="center">
 
	<table border="0" width="640" id="new_schedule" bgcolor="#CFD7DB">
		<tr>
			<td  width="128"><font size="2">New Schedule:</font></td>
			<td>&nbsp;</td>
			<td width="493">
			<input  type="text" name="new_name" style="width: 400" size="40" >
			
			</td>
		</tr>
		<tr>
			<td><font size="2">Schedule Type:</font></td>
			<td>&nbsp;</td>
			<td><select size="1" name="schedule_type">
			<option selected value="4">Max./Min.</option>
			</select></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td> <input type="submit" value="  Next &gt;&gt;  " name="B1"></td>
		</tr>
	</table>
	</div>
  <p align="center">&nbsp;</p>
	<input type="hidden" name="car_rate_rule_schedule_id" value="0">
</form>


<p align="center">&nbsp;</p>
<div align="center">
	<table border="0" width="640" id="table3">
		<tr>
			<td><font size="2"><b>Directions:</b> Either select an existing 
			schedule from the list above or select &quot;New Schedule&quot; from the 
			drop-down list to create a new rule schedule. If you select a new 
			schedule you will be able to create a descriptive name for it in the 
			next step</font></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
</div>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>