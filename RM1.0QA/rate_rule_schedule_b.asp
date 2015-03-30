<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 360

	On Error Resume Next

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim intScheduleID
	Dim strType

	intScheduleID = Request.Form("car_rate_rule_schedule_id")
	strType = Request("type")
	
	If strType = "new" Then
		intScheduleID = 0
	
	Else
		If IsNumeric(intScheduleID ) Then

			strConn = Session("pro_con")
	
		  	Set adoRS = CreateObject("ADODB.Recordset")
		  	Set adoCmd = CreateObject("ADODB.Command")
		
			adoCmd.ActiveConnection = strConn
			adoCmd.CommandText = "car_rate_rule_schedule_detail_1_select"
			adoCmd.CommandType = 4
	
			adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_schedule_id", 3, 1, 0, intScheduleID)

			Set adoRS = adoCmd.Execute
	
		End If

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
	End If


	%>

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor by Rate-Highway, Inc. | Rate Rule Schedule Management</title>
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
<p align="center">
<% Dim intScheduleType 
   If strType = "new" Then
     intScheduleType = Request.Form("schedule_type")
   Else
   	 intScheduleType = adoRS.Fields("schedule_type_id").Value 
   End If

%> 
<% Select Case intScheduleType %>
<%   Case 1                    %>
<font size="5" color="#384F5B">Days Out / Utilization Schedule (Rate Amount Diff.)</font></p>
<%   Case 2                    %>
<font size="5" color="#384F5B">Days Out / Utilization Schedule (Rate Maximum)</font></p>
<%   Case 3                    %>
<font size="5" color="#384F5B">Days Out / Utilization Schedule (Rate Minimum)</font></p>
<% End Select                  %>
<form method="POST" action="rate_rule_schedule_update.asp" name="rate_rule_schedule">
  <div align="center">
	<table border="0" width="640" id="table2">
		<tr>
			<td width="113"><font size="2">Schedule Name: </font></td>
			<td>&nbsp;</td>
			<td>
			<% If strType = "new" Then  %>
			<input type="text" name="name" size="60" value="<%=Request.Form("new_name") %>" ></td>
			<% Else   %>
			<input type="text" name="name" size="60" value="<%=adoRS.Fields("car_rate_rule_schedule_desc").Value %>" ></td>
			<% End If %>
		</tr>
		<tr>
			<td width="113">&nbsp;</td>
			<td>&nbsp;</td>
			<% If strType = "new" Then  %>
			<td><input type="checkbox" name="save_copy" value="TRUE" id="save_copy" disabled><label for="save_copy">Save as a copy</label></td>
			<% Else   %>
			<td><input type="checkbox" name="save_copy" value="TRUE" id="save_copy"><label for="save_copy">Save	as a copy</label></td>
			<% End If %>

		</tr>
	</table>
	<table border="0" width="600" id="tblSchedule" name="tblSchedule">
		<tr>
			<td bgcolor="#384F5B" bordercolor="#CFD7DB" colspan="11">
			<p align="center"><font color="#FFFFFF">Days Out<font size="1"><br>
			(number indicates number of days cap)</font></font></td>
		</tr>
		<tr>
			<td bgcolor="#384F5B" bordercolor="#384F5B" rowspan="12">
			<font size="2" color="#FFFFFF">&nbsp;Util. </font>
			<p><font size="2" color="#FFFFFF">&nbsp;Min</font><p align="center">
			<font size="2" color="#FFFFFF">%</font></td>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">&nbsp;</td>
			<td bgcolor="#CFD7DB" style="width: 75">
			<input type="text" name="days_out_grp1" size="10" value="Same" class="header"></td>
			<td bgcolor="#CFD7DB">
			<input type="text" name="days_out_grp2" size="10" value="Next" class="header"></td>
			<td bgcolor="#CFD7DB">
			<input type="text" name="days_out_grp3" size="10" value="4" class="header"></td>
			<td bgcolor="#CFD7DB">
			<input type="text" name="days_out_grp4" size="10" value="7" class="header"></td>
			<td bgcolor="#CFD7DB">
			<input type="text" name="days_out_grp5" size="10" value="14" class="header"></td>
			<td bgcolor="#CFD7DB">
			<input type="text" name="days_out_grp6" size="10" value="30" class="header"></td>
			<td bgcolor="#CFD7DB">
			<input type="text" name="days_out_grp7" size="10" value="50" class="header"></td>
			<td bgcolor="#CFD7DB">
			<input type="text" name="days_out_grp8" size="10" value="+" class="header"></td>
			<td bgcolor="#CFD7DB">
			<input type="text" name="days_out_grp9" size="10" value=""  class="header"></td>
		</tr>
		<% 
		   Dim intCol                 
		   Dim intCol0                
		   Dim intCol1                
		   Dim intCol2                
		   Dim intCol3                
		   Dim intCol4                
		   Dim intCol5                
		   Dim intCol6                
		   Dim intCol7                
		   Dim intCol8                
		   Dim intCol9                
		   Dim intLast                
		   Dim intRowCount                 

		   If strType = "existing" Then

			   Set adoRS = adoRS.NextRecordset

			   While adoRS.EOF = False   
			      intCol = adoRS.Fields("util_min").Value   
			      intLast = adoRS.Fields("util_min").Value   
		
		
			   intCol0 = ""               
			   intCol1 = ""               
			   intCol2 = ""               
			   intCol3 = ""               
			   intCol4 = ""               
			   intCol5 = ""               
			   intCol6 = ""               
			   intCol7 = ""               
			   intCol8 = ""               
			   intCol9 = ""               

				
		   
		
				While (intCol = intLast) And (adoRS.EOF = False)
				
					Select Case adoRS.Fields("days_out_grp").Value
						Case 0
							intCol0 = adoRS.Fields("response_amt").Value
						Case 1
							intCol1 = adoRS.Fields("response_amt").Value
						Case 2
							intCol2 = adoRS.Fields("response_amt").Value
						Case 3
							intCol3 = adoRS.Fields("response_amt").Value
						Case 4
							intCol4 = adoRS.Fields("response_amt").Value
						Case 5
							intCol5 = adoRS.Fields("response_amt").Value
						Case 6
							intCol6 = adoRS.Fields("response_amt").Value
						Case 7
							intCol7 = adoRS.Fields("response_amt").Value
						Case 8
							intCol8 = adoRS.Fields("response_amt").Value
						Case 9
							intCol9 = adoRS.Fields("response_amt").Value
							
					End Select

					adoRS.MoveNext
				
					If (adoRS.EOF = False) Then
						intCol = adoRS.Fields("util_min").Value
					Else
						intCol = 999999
					End If 

				Wend

	
		
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
		
		
		<% 'If adoRS.EOF = False Then       %>
		<%   intRowCount = intRowCount + 1 %>
		
		
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell<%=intRowCount %>" size="10" style="background-color: #CCCCCC" class="data_cell" value="<%=intLast %>"></td>
			<td>
			<input type="text" name="cell<%=intRowCount %>" size="10" class="data_cell" value="<%=intCol0 %>"></td>
			<td>
			<input type="text" name="cell<%=intRowCount %>" size="10" class="data_cell" value="<%=intCol1 %>"></td>
			<td>
			<input type="text" name="cell<%=intRowCount %>" size="10"  class="data_cell" value="<%=intCol2 %>"></td>
			<td>
			<input type="text" name="cell<%=intRowCount %>" size="10"  class="data_cell" value="<%=intCol3 %>"></td>
			<td>
			<input type="text" name="cell<%=intRowCount %>" size="10"  class="data_cell" value="<%=intCol4 %>"></td>
			<td>
			<input type="text" name="cell<%=intRowCount %>" size="10"  class="data_cell" value="<%=intCol5 %>"></td>
			<td>
			<input type="text" name="cell<%=intRowCount %>" size="10"  class="data_cell" value="<%=intCol6 %>"></td>
			<td>
			<input type="text" name="cell<%=intRowCount %>" size="10" class="data_cell" value="<%=intCol7 %>"></td>
			<td>
			<input type="text" name="cell<%=intRowCount %>" size="10" class="data_cell" value="<%=intCol8 %>"></td>
		</tr>
		
		<% 'End If %>
		
		<% Wend              %>
		
		<% End If %>
					
		<% If strType = "new" Then %>
		
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell1" size="10"  class="data_cell" value="40" style="background-color: #CCCCCC"></td>
			<td>
			<input type="text" name="cell1" size="10"  class="data_cell"></td>
			<td>
			<input type="text" name="cell1" size="10"  class="data_cell"></td>
			<td>
			<input type="text" name="cell1" size="10"  class="data_cell"></td>
			<td>
			<input type="text" name="cell1" size="10"  class="data_cell"></td>
			<td>
			<input type="text" name="cell1" size="10"  class="data_cell"></td>
			<td>
			<input type="text" name="cell1" size="10"  class="data_cell"></td>
			<td>
			<input type="text" name="cell1" size="10"  class="data_cell"></td>
			<td>
			<input type="text" name="cell1" size="10"  class="data_cell"></td>
			<td>
			<input type="text" name="cell1" size="10"  class="data_cell"></td>
		</tr>
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell2" size="10" value="46"  style="background-color: #CCCCCC" class="data_cell"></td>
			<td>
			<input type="text" name="cell2" size="10" class="data_cell"></td>
			<td>
			<input type="text" name="cell2" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell2" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell2" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell2" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell2" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell2" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell2" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell2" size="10" class="data_cell"></td>
		</tr>
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell3" size="10" value="51"  style="background-color: #CCCCCC" class="data_cell"></td>
			<td>
			<input type="text" name="cell3" size="10" class="data_cell"></td>
			<td>
			<input type="text" name="cell3" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell3" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell3" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell3" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell3" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell3" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell3" size="10"  class="data_cell" ></td>
			<td>
			<input type="text" name="cell3" size="10" class="data_cell"></td>
		</tr>
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell4" size="10" value="56"  style="background-color: #CCCCCC" class="data_cell"></td>
			<td><input type="text" name="cell4" size="10" class="data_cell"></td>
			<td><input type="text" name="cell4" size="10" class="data_cell"></td>
			<td><input type="text" name="cell4" size="10" class="data_cell"></td>
			<td><input type="text" name="cell4" size="10" class="data_cell" ></td>
			<td><input type="text" name="cell4" size="10" class="data_cell" ></td>
			<td><input type="text" name="cell4" size="10" class="data_cell" ></td>
			<td><input type="text" name="cell4" size="10" class="data_cell"></td>
			<td><input type="text" name="cell4" size="10" class="data_cell"></td>
			<td><input type="text" name="cell4" size="10" class="data_cell"></td>
		</tr>
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell5" size="10" value="61"  style="background-color: #CCCCCC" class="data_cell"></td>
			<td><input type="text" name="cell5" size="10" class="data_cell"></td>
			<td><input type="text" name="cell5" size="10" class="data_cell"></td>
			<td><input type="text" name="cell5" size="10" class="data_cell"></td>
			<td><input type="text" name="cell5" size="10" class="data_cell"></td>
			<td><input type="text" name="cell5" size="10" class="data_cell"></td>
			<td><input type="text" name="cell5" size="10" class="data_cell"></td>
			<td><input type="text" name="cell5" size="10" class="data_cell"></td>
			<td><input type="text" name="cell5" size="10" class="data_cell"></td>
			<td><input type="text" name="cell5" size="10" class="data_cell"></td>
		</tr>
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell6" size="10" value="66"  style="background-color: #CCCCCC" class="data_cell"></td>
			<td><input type="text" name="cell6" size="10" class="data_cell"></td>
			<td><input type="text" name="cell6" size="10" class="data_cell"></td>
			<td><input type="text" name="cell6" size="10" class="data_cell"></td>
			<td><input type="text" name="cell6" size="10" class="data_cell"></td>
			<td><input type="text" name="cell6" size="10" class="data_cell"></td>
			<td><input type="text" name="cell6" size="10" class="data_cell"></td>
			<td><input type="text" name="cell6" size="10" class="data_cell"></td>
			<td><input type="text" name="cell6" size="10" class="data_cell"></td>
			<td><input type="text" name="cell6" size="10" class="data_cell"></td>
		</tr>
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell7" size="10"  value="71" style="background-color: #CCCCCC" class="data_cell"></td>
			<td><input type="text" name="cell7" size="10"class="data_cell"></td>
			<td><input type="text" name="cell7" size="10" class="data_cell"></td>
			<td><input type="text" name="cell7" size="10" class="data_cell"></td>
			<td><input type="text" name="cell7" size="10" class="data_cell"></td>
			<td><input type="text" name="cell7" size="10" class="data_cell"></td>
			<td><input type="text" name="cell7" size="10" class="data_cell"></td>
			<td><input type="text" name="cell7" size="10" class="data_cell"></td>
			<td><input type="text" name="cell7" size="10" class="data_cell"></td>
			<td><input type="text" name="cell7" size="10" class="data_cell"></td>
		</tr>
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell8" size="10"  value="76" style="background-color: #CCCCCC" class="data_cell"></td>
			<td><input type="text" name="cell8" size="10" class="data_cell"></td>
			<td><input type="text" name="cell8" size="10" class="data_cell"></td>
			<td><input type="text" name="cell8" size="10" class="data_cell"></td>
			<td><input type="text" name="cell8" size="10" class="data_cell"></td>
			<td><input type="text" name="cell8" size="10" class="data_cell"></td>
			<td><input type="text" name="cell8" size="10" class="data_cell"></td>
			<td><input type="text" name="cell8" size="10" class="data_cell"></td>
			<td><input type="text" name="cell8" size="10" class="data_cell"></td>
			<td><input type="text" name="cell8" size="10" class="data_cell"></td>
		</tr>
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell9" size="10"  style="background-color: #CCCCCC" class="data_cell"></td>
			<td><input type="text" name="cell9" size="10" class="data_cell"></td>
			<td><input type="text" name="cell9" size="10" class="data_cell"></td>
			<td><input type="text" name="cell9" size="10" class="data_cell"></td>
			<td><input type="text" name="cell9" size="10" class="data_cell"></td>
			<td><input type="text" name="cell9" size="10" class="data_cell"></td>
			<td><input type="text" name="cell9" size="10" class="data_cell"></td>
			<td><input type="text" name="cell9" size="10" class="data_cell"></td>
			<td><input type="text" name="cell9" size="10" class="data_cell"></td>
			<td><input type="text" name="cell9" size="10" class="data_cell"></td>
		</tr>
		<tr>
			<td bgcolor="#CFD7DB" bordercolor="#CFD7DB">
			<input type="text" name="cell10" size="10"  style="background-color: #CCCCCC" class="data_cell"></td>
			<td><input type="text" name="cell10" size="10" class="data_cell"></td>
			<td><input type="text" name="cell10" size="10" class="data_cell"></td>
			<td><input type="text" name="cell10" size="10" class="data_cell"></td>
			<td><input type="text" name="cell10" size="10" class="data_cell"></td>
			<td><input type="text" name="cell10" size="10" class="data_cell"></td>
			<td><input type="text" name="cell10" size="10" class="data_cell"></td>
			<td><input type="text" name="cell10" size="10" class="data_cell"></td>
			<td><input type="text" name="cell10" size="10" class="data_cell"></td>
			<td><input type="text" name="cell10" size="10" class="data_cell"></td>
		</tr>
		
		<% End If %>
		
	</table>
	<p><font size="2">&nbsp;<a href="javascript:addRowToTable()">add row</a> |
	</font>
	<a href="javascript:addRowToTable()"><font size="2">add column</font></a></p>
	<p><input type="submit" value="  Save  " name="submit"></div>
  <p align="center">&nbsp;</p>
	<input type="hidden" name="schedule_type" value="<%=intScheduleType %>">
	<input type="hidden" name="schedule_id" value="<%=intScheduleID %>">
</form>
<p align="center">&nbsp;</p>
<p align="center"><font size="2">Directions: Enter the information in the cells 
for this schedule. To enter a positive number (+) simply enter<br>
the number without any sign, for a negative number (-) enter a 
minus/dash in front of the number. <%=intScheduleID %></p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>