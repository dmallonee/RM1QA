<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 60

	On Error Resume Next

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim intScheduleID
	Dim strType
	Dim strUserId
	Dim intRowCount
	
	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	intScheduleID = Request.Form("car_rate_rule_schedule_id")
	strCityCd = Request.Form("city_cd")
	If Len(strCityCd) = 0 Then
		strCityCd = "***"
	End If
	
	intMonth = CInt(Request.Form("month"))
	If (intMonth < 0) Or (intMonth > 12) Then
		intMonth = 0
	End If
	
	strType = Request("type")
	
	strConn = Session("pro_con")
	
	If strType = "new" Then
		intScheduleID = 0
	
	Else

		If IsNumeric(intScheduleID) Then

			strConn = Session("pro_con")
	
		  	Set adoRS = CreateObject("ADODB.Recordset")
		  	Set adoCmd = CreateObject("ADODB.Command")
			
			adoCmd.ActiveConnection = strConn
			adoCmd.CommandText = "car_rate_rule_schedule_detail_4_select"
			adoCmd.CommandType = 4
		
			adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_schedule_id",   3, 1, 0, intScheduleID)
			adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",                   200, 1, 6, "***") 'strCityCd)
			adoCmd.Parameters.Append adoCmd.CreateParameter("@month",                      17, 1, 0, 0) 'intMonth)
	
			Set adoRS = adoCmd.Execute
	
		End If

	End If	
	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Error while collecting schedule detail!<br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If


	Rem Get the car types
	Set adoCmdCarTypes = CreateObject("ADODB.Command")

	adoCmdCarTypes.ActiveConnection =  strConn
	adoCmdCarTypes.CommandText = "car_type_select"
	adoCmdCarTypes.CommandType = 4
	
	adoCmdCarTypes.Parameters.Append adoCmdCarTypes.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRSCarTypes = adoCmdCarTypes.Execute

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Error occured while collecting car types<br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If

	Rem Get the car types
	Set adoCmdCities = CreateObject("ADODB.Command")

	adoCmdCities.ActiveConnection =  strConn
	adoCmdCities.CommandText = "user_city_select"
	adoCmdCities.CommandType = 4
	
	adoCmdCities.Parameters.Append adoCmdCities.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRSCities = adoCmdCities.Execute

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Error occured while collecting cities<br>"
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
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor by Rate-Highway, Inc. | Rate Rule Schedule Management</title>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script language="javascript" type="text/javascript" src="inc/sitewide.js" ></script>
<script language="javascript" type="text/javascript"  src="inc/header_menu_support.js"></script>
<script language="javascript" type="text/javascript" > 
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

// ADD and REMOVE row support
function addRowToTable()
{
  var tbl = document.getElementById('tblSchedule');
  var lastRow = tbl.rows.length;
  // if there's no header row in the table, then iteration = lastRow + 1
  var iteration = lastRow;
  var row = tbl.insertRow(lastRow);
  
  row.bgcolor = '#CFD7DB';
  row.bordercolor = '#CFD7DB';
  
 
  // left cell
  var cellLeft = row.insertCell(0);
  var textNode = document.createElement('input');
  textNode.size = 10;

  //textNode.style = 'data_cell';
  cellLeft.appendChild(textNode);
  
  // right cell
  var cellRight = row.insertCell(1);
  var el = document.createElement('input');
  el.type = 'input';
  el.name = 'txtRow' + iteration;
  el.id = 'txtRow' + iteration;
  el.size = 10;
  
//  el.onkeypress = keyPressTest;
  cellRight.appendChild(el);
  
  // select cell
  var cellRight = row.insertCell(2);
  var el = document.createElement('input');
  el.type = 'input';
  el.name = 'txtRow' + iteration;
  el.id = 'txtRow' + iteration;
  el.size = 10;
  cellRight.appendChild(el);

}
function keyPressTest(e, obj)
{
  var validateChkb = document.getElementById('chkValidateOnKeyPress');
  if (validateChkb.checked) {
    var displayObj = document.getElementById('spanOutput');
    var key;
    if(window.event) {
      key = window.event.keyCode; 
    }
    else if(e.which) {
      key = e.which;
    }
    var objId;
    if (obj != null) {
      objId = obj.id;
    } else {
      objId = this.id;
    }
    displayObj.innerHTML = objId + ' : ' + String.fromCharCode(key);
  }
}
function removeRowFromTable()
{
  var tbl = document.getElementById('tblSchedule');
  var lastRow = tbl.rows.length;
  if (lastRow > 2) tbl.deleteRow(lastRow - 1);
}
function openInNewWindow(frm)
{
  // open a blank window
  var aWindow = window.open('', 'TableAddRowNewWindow',
   'scrollbars=yes,menubar=yes,resizable=yes,toolbar=no,width=400,height=400');
   
  // set the target to the blank window
  frm.target = 'TableAddRowNewWindow';
  
  // submit
  frm.submit();
}
function validateRow(frm)
{
  var chkb = document.getElementById('chkValidate');
  if (chkb.checked) {
    var tbl = document.getElementById('tblSample');
    var lastRow = tbl.rows.length - 1;
    var i;
    for (i=1; i<=lastRow; i++) {
      var aRow = document.getElementById('txtRow' + i);
      if (aRow.value.length <= 0) {
        alert('Row ' + i + ' is empty');
        return;
      }
    }
  }
  openInNewWindow(frm);
}



function disableFields2()
{
	document.rate_rule_schedule.min_1.disabled=true;
	document.rate_rule_schedule.max_1.disabled=true;
}	
	
function disableFields()
{
	for(i=0; i<document.rate_rule_schedule.elements.length; i++)
	{
	
		if(document.rate_rule_schedule.elements[i].class=="data_cell")
		{
			document.rate_rule_schedule.elements[i].disabled=true;
		}
	}
}	


</script> 

<style type="text/css" >
<!--
P {
	COLOR: navy; FONT-FAMILY: Verdana, Arial, sans-serif
}
.data_cell   { width: 65; text-align: right; font-family: Tahoma; font-size: 10pt }
.data_cell_ctr   { width: 65; text-align: center; font-family: Tahoma; font-size: 10pt }
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
<p align="center">
<% Dim intScheduleType 
   If strType = "new" Then
     intScheduleType = Request.Form("schedule_type")
   Else
   	 intScheduleType = adoRS.Fields("schedule_type_id").Value 
   End If

%>  
<% Select Case intScheduleType %>
<%   Case 4                    %>
<font size="5" color="#384F5B">Rule Max. / Min. Schedule</font></p>
<% End Select                  %>
  <div align="center">
  <form method="POST" action="rate_rule_maxmin_schedule_b.asp" name="rate_rule_schedule_redisplay">
	<table border="0" width="640" id="header">
		<tr>
			<td width="113"><font size="2">Schedule Name: </font></td>
			<td>&nbsp;</td>
			<td colspan="2">
			<% 
			   If strType = "new" Then  
			     strName = Request.Form("new_name") 
			   Else   
			     strName = adoRS.Fields("car_rate_rule_schedule_desc").Value 
			
			   End If 
			%>
			<input type="text" name="name" size="60" value="<%=strName %>" tabindex="1" >

			</td>
		</tr>
		<tr>
			<td width="113">&nbsp;</td>
			<td>&nbsp;</td>
			<% If strType = "new" Then  %>
			<td colspan="2">
			<input type="checkbox" name="save_copy" value="TRUE" id="save_copy" disabled tabindex="2"><label for="save_copy">Save as a copy</label></td>
			<% Else   %>
			<td>
			<input type="checkbox" name="save_copy" value="TRUE" id="save_copy" tabindex="2"><label for="save_copy">Save	as a copy</label></td>
			<% End If %>

		</tr>
		<tr>
			<td width="113"><font size="2">City:</font></td>
			<td>&nbsp;</td>
			<%
			
			If strCityCd = "***" Then
				strCityCdDesc = "Default"
			Else
				strCityCdDesc = strCityCd
			End If
			
			
			Select Case intMonth
			
				Case 0
					strMonth = "Default"
				Case 1
					strMonth = "Jan"
				Case 2
					strMonth = "Feb"
				Case 3
					strMonth = "Mar"
				Case 4
					strMonth = "Apr"
				Case 5
					strMonth = "Default"
				Case 6
					strMonth = "Default"
				Case 7
					strMonth = "Default"
				Case 8
					strMonth = "Default"
				Case 9
					strMonth = "Default"
				Case 10
					strMonth = "Default"
				Case 11
					strMonth = "Default"
				Case 12
					strMonth = "Default"
					
			
			End Select			
			
			%>
			<td><input name="city_cd_copy" type="text" value="<%=strCityCdDesc %>" readonly="readonly" size="8" tabindex="3"></td>
			<td><select name="city_cd" tabindex="5" onchange="disableFields()">
			<option selected="" value="***">Default</option>

			<% While (adoRSCities.EOF = False) %>
				<% If strCityCd = adoRSCities.Fields("city_cd").Value Then %>
				<option selected="" value="<%=adoRSCities.Fields("city_cd").Value %>"><%=adoRSCities.Fields("city_cd").Value %></option>
				<% Else %>
				<option  value="<%=adoRSCities.Fields("city_cd").Value %>"><%=adoRSCities.Fields("city_cd").Value %></option>				
				<% End If %>
				<% adoRSCities.MoveNext %>
			<% Wend %>
			
			</select></td>
			<td>&nbsp;</td>

		</tr>
		<tr>
			<td width="113"><font size="2">Month:</font></td>
			<td>&nbsp;</td>
			<td>
			<input name="month_copy" type="text" value="<%=strMonth %>" readonly="readonly" size="8" tabindex="4"></td>
			<td>
			<select name="month" tabindex="6">

			<% If intMonth = 0 Then %>
			<option selected="" value="0">Default</option>
			<% Else %>
			<option value="0">Default</option>
			<% End If %>

			<% If intMonth = 1 Then %>
			<option selected="" value="1">Jan</option>
			<% Else %>
			<option value="1">Jan</option>
			<% End If %>

			<% If intMonth = 2 Then %>
			<option selected="" value="2">Feb</option>
			<% Else %>
			<option value="2">Feb</option>
			<% End If %>
			
			<% If intMonth = 3 Then %>
			<option selected="" value="3">Mar</option>
			<% Else %>
			<option value="3">Mar</option>
			<% End If %>

			<% If intMonth = 4 Then %>
			<option selected="" value="4">Apr</option>
			<% Else %>
			<option value="4">Apr</option>
			<% End If %>

			<% If intMonth = 5 Then %>
			<option selected="" value="5">May</option>
			<% Else %>
			<option value="5">May</option>
			<% End If %>

			<% If intMonth = 6 Then %>
			<option selected="" value="6">Jun</option>
			<% Else %>
			<option value="6">Jun</option>
			<% End If %>
			
			<% If intMonth = 7 Then %>
			<option selected="" value="7">Jul</option>
			<% Else %>
			<option value="7">Jul</option>
			<% End If %>

			<% If intMonth = 8 Then %>
			<option selected="" value="8">Aug</option>
			<% Else %>
			<option value="8">Aug</option>
			<% End If %>

			<% If intMonth = 9 Then %>
			<option selected="" value="9">Sep</option>
			<% Else %>
			<option value="9">Sep</option>
			<% End If %>

			<% If intMonth = 10 Then %>
			<option selected="" value="10">Oct</option>
			<% Else %>
			<option value="10">Oct</option>
			<% End If %>

			<% If intMonth = 11 Then %>
			<option selected="" value="11">Nov</option>
			<% Else %>
			<option value="11">Nov</option>
			<% End If %>

			<% If intMonth = 12 Then %>
			<option selected="" value="12">Dec</option>
			<% Else %>
			<option value="12">Dec</option>
			<% End If %>

			</select></td>
			<td>&nbsp;</td>

		</tr>
		<tr>
			<td width="113">&nbsp;</td>
			<td>&nbsp;</td>
			<td>
			&nbsp;</td>
			<td>
			<input name="display" type="submit" value="Display" tabindex="7"></td>
			<td>&nbsp;</td>

		</tr>
		<tr>
			<td width="113">&nbsp;</td>
			<td>&nbsp;</td>
			<td colspan="2">&nbsp;</td>
			<td>&nbsp;</td>

		</tr>
	</table>
	  <input type="hidden" name="schedule_type" value="<%=intScheduleType %>">
	  <input name="type" type="hidden" value="existing">
	  <input name="car_rate_rule_schedule_id" type="hidden" value="<%=intScheduleID %>">
	</form>
 </div>
 <div align="center" >
	<form method="POST" action="rate_rule_maxmin_schedule_update.asp" name="rate_rule_schedule">

	<table border="0" width="250" id="tblSchedule" name="tblSchedule">
		<tr>
			<td bgcolor="#384F5B" bordercolor="#CFD7DB" colspan="4">
			<p align="center"><font color="#FFFFFF">Rate Amounts</font><font size="1" color="#FFFFFF"><br>
			(Empty cells mean no limit)</font></td>
		</tr>
		<tr>
			<td bgcolor="#384F5B" bordercolor="#384F5B" rowspan="100" align="center" width="60" style="text-align: center">
			<font size="2" color="#FFFFFF">&nbsp;Car <br>
			&nbsp;Type</font></td>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">&nbsp;</td>
			<td bgcolor="#CFD7DB" style="width: 75">
			<input type="text" name="days_out_grp1" size="10" value="Min." class="header"></td>
			<td bgcolor="#CFD7DB">
			<input type="text" name="days_out_grp2" size="10" value="Max." class="header"></td>
		</tr>
		
		<% intRowCount = 0 %>

		<% If strType = "new" Then %>
			<% While (adoRSCarTypes.EOF = False) %>
				<% intRowCount = intRowCount + 1 %>
				<tr>
				<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
				<input type="text" name="cell<%=intRowCount %>" size="10" style="background-color: #CCCCCC" class="data_cell_ctr" value="<%=adoRSCarTypes.Fields("car_type_cd").Value %>"></td>
				<td>
				<input type="text" name="cell<%=intRowCount %>" size="10" class="data_cell" value="" id="min_<%=intRowCount %>"></td>
				<td>
				<input type="text" name="cell<%=intRowCount %>" size="10" class="data_cell" value="" id="max_<%=intRowCount %>" ></td>
				</tr>
				<% adoRSCarTypes.MoveNext %>
			<% Wend %>
		<% Else %>
			<% Set adoRS = adoRS.NextRecordset %>
			<% While (adoRS.EOF = False) %>
				<% intRowCount = intRowCount + 1 %>
				<tr>
				<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
				<input type="text" name="cell<%=intRowCount %>" size="10" style="background-color: #CCCCCC" class="data_cell_ctr" value="<%=adoRS.Fields("car_type_cd").Value %>"></td>
				<td>
				<input type="text" name="cell<%=intRowCount %>" size="10" class="data_cell" value="<%=adoRS.Fields("min_amt").Value  %>" id="min_<%=intRowCount %>"></td>
				<td>
				<input type="text" name="cell<%=intRowCount %>" size="10" class="data_cell" value="<%=adoRS.Fields("max_amt").Value  %>" id="max_<%=intRowCount %>"></td>
				</tr>
				<% adoRS.MoveNext %>
			<% Wend %>
	
	
	
		<% End If %>	

<%
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Error occured while creating table<br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If

%>		
		
	</table>
	<p><input type="submit" value="  Save  " name="submit" tabindex="8"></p>
  <p align="center">&nbsp;</p>
	<input type="hidden" name="schedule_type" value="<%=intScheduleType %>">
	<input type="hidden" name="schedule_id" value="<%=intScheduleID %>">
		<input name="city_cd" type="hidden" value="<%=strCityCd %>">
		<input name="month" type="hidden" value="<%=intMonth %>">
		<input name="name" type="hidden" value="<%=strName  %>">
</form>
</div>
<div align="center">
	<table border="0" width="500" id="table3">
		<tr>
			<td><font size="2">Directions: Enter the rate amount minimums in the 
			min field, and enter the rate amount maximum that you want the rules 
			to observe in the max field. Do this for each car type. These fields 
			are strictly optional and if you decide not to enter values your 
			rules will still process; however, you will not have your rates 
			limited by the minimums and maximums set here.</font>


			<p>&nbsp;</p></td>
		</tr>
		<tr>
			<td><font size="2">Note, if you do not see a car type here that you 
			need in your account, please contact Customer Support at
			<a href="mailto:support@rate-highway.com">support@rate-highway.com</a> 
			and let them know the car type(s) you would like added to your 
			account. </font></td>
			<td>
		</tr>
		<tr>
			<td>
			<font size="1">
			<%=intScheduleID %><br>
			<%=strCityCd %><br>
			<%=intMonth %><br>
			<%=intScheduleType %><br>
			<%=intScheduleID %>
			</font>
			</td>
		</tr>
	</table>
</div>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>