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
		
			adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_schedule_id", 3, 1, 0, intScheduleID)
	
			'Set adoRS = adoCmd.Execute
	
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
		
	'Set adoRSCarTypes = adoCmdCarTypes.Execute

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Error occured while collecting car types<br>"
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
<script type="text/javascript" language="javascript" src="inc/sitewide.js" ></script>
<script type="text/javascript" language="javascript" src="inc/header_menu_support.js"></script>
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

td {
	font-size: x-small;
	font-weight: normal;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}

.style1 {
	text-align: center;
}

.style2 {
	background-color: #F0F0EA;
}

-->
</style>
</head>
<body topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_middle.jpg"><img src="images/top_left.jpg" width="423" height="91"></td>
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
   	 intScheduleType = 4 'adoRS.Fields("schedule_type_id").Value 
   End If

%>  
<font size="5" color="#384F5B">Search Schedule</font></p>
<form method="POST" action="profile_search_schedule_update.asp" name="profile_search_schedule">
	<table border="0" width="640" id="new_profile" align="center" cellspacing="0" cellpadding="0">
		<tr>
			<td width="113"><font size="2">&nbsp;Schedule Name: </font></td>
			<td>&nbsp;</td>
			<td>
			<% If strType = "new" Then  %>
			<input type="text" name="name" size="60" value="<%=Request.Form("new_name") %>" >
			<% Else   %>
			<!-- adoRS.Fields("car_rate_rule_schedule_desc").Value -->
			<input type="text" name="name" size="60" value="test" >
			<% End If %>
			</td>
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
		<tr>
			<td width="113">&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>

		</tr>
		<tr>
		<td class="style2">&nbsp;</td>
		<td width="162" class="style2">
		&nbsp;</td>
                    <td width="365" class="style2">
                  
                    &nbsp;</td>
                    <td class="style2">&nbsp;</td>
                  </tr>
		<tr>
		<td class="style2">&nbsp; Time:</td>
		<td width="162" class="style2">
		<input name="time" type="radio" checked="checked" value="fixed">Fixed: </td>
                    <td width="365" class="style2">
                  
                    <select name="scheduled_time" style="width:175" size="1">
                
                <% If InStr(1, strCheckTime, "00:00") > 0 Then %>
                	 <option selected value='00:00'>Midnight</option>
                
                <% Else %>
                	 <option value='00:00'>Midnight</option>
                
                <% End If %>
                
                <%

                   For intIndex = 1 To 11  
                     For intMinuteIndex = 0 To 3 
                     Select Case intMinuteIndex
                     	Case 0
	                       strTime = intIndex & ":00 am"
                     	Case 1
	                       strTime = intIndex & ":15 am"
                     	Case 2
	                       strTime = intIndex & ":30 am"
                     	Case 3
	                       strTime = intIndex & ":45 am"
					 End Select	                       
                     strTimeValue = trim(FormatDateTime(strTime, 4))
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
					 Next
                   Next   


				   If strCheckTime = "12:00" Then	
                %>
                	 <option selected value='12:00'>Noon</option>
                <%
				   Else
                %>
                	 <option value='12:00'>Noon</option>
                <%
				   End If



                   For intIndex = 1 To 11 
                     For intMinuteIndex = 0 To 3 
                     Select Case intMinuteIndex
                     	Case 0
	                       strTime = intIndex & ":00 pm"
                     	Case 1
	                       strTime = intIndex & ":15 pm"
                     	Case 2
	                       strTime = intIndex & ":30 pm"
                     	Case 3
	                       strTime = intIndex & ":45 pm"
					 End Select	                       
                     strTimeValue = trim(FormatDateTime(strTime, 4))
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                     
   				   	 If intIndex <> 8 Then	
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
					 End If

				     Next
                   Next   
                %>
                    </select> (local time).</td>
                    <td class="style2"></td>
                  </tr>
                  <tr>
                  	<td class="style2"></td>
                  	<td class="style2">&nbsp;&nbsp;&nbsp;&nbsp;or</td>
                  	<td class="style2">&nbsp;</td>
                  	<td class="style2"></td>
                  </tr>
                  <tr>
                  	<td class="style2"></td>
                  	<td class="style2"><input name="time" type="radio" value="random">Random:</td>
                  	<td class="style2">Between:
                  
                    <select name="scheduled_time0" style="width:100" size="1">
                
                <% If InStr(1, strCheckTime, "00:00") > 0 Then %>
                	 <option selected value='00:00'>Midnight</option>
                
                <% Else %>
                	 <option value='00:00'>Midnight</option>
                
                <% End If %>
                
                <%

                   For intIndex = 1 To 11  
                     For intMinuteIndex = 0 To 3 
                     Select Case intMinuteIndex
                     	Case 0
	                       strTime = intIndex & ":00 am"
                     	Case 1
	                       strTime = intIndex & ":15 am"
                     	Case 2
	                       strTime = intIndex & ":30 am"
                     	Case 3
	                       strTime = intIndex & ":45 am"
					 End Select	                       
                     strTimeValue = trim(FormatDateTime(strTime, 4))
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
					 Next
                   Next   


				   If strCheckTime = "12:00" Then	
                %>
                	 <option selected value='12:00'>Noon</option>
                <%
				   Else
                %>
                	 <option value='12:00'>Noon</option>
                <%
				   End If



                   For intIndex = 1 To 11 
                     For intMinuteIndex = 0 To 3 
                     Select Case intMinuteIndex
                     	Case 0
	                       strTime = intIndex & ":00 pm"
                     	Case 1
	                       strTime = intIndex & ":15 pm"
                     	Case 2
	                       strTime = intIndex & ":30 pm"
                     	Case 3
	                       strTime = intIndex & ":45 pm"
					 End Select	                       
                     strTimeValue = trim(FormatDateTime(strTime, 4))
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                     
   				   	 If intIndex <> 8 Then	
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
					 End If

				     Next
                   Next   
                %>
                    </select>&nbsp; and&nbsp;
                  
                    <select name="scheduled_time1" style="width:100" size="1">
               
                <% If InStr(1, strCheckTime, "00:00") > 0 Then %>
                	 <option selected value='00:00'>Midnight</option>
                
                <% Else %>
                	 <option value='00:00'>Midnight</option>
                
                <% End If %>
                
                <%

                   For intIndex = 1 To 11  
                     For intMinuteIndex = 0 To 3 
                     Select Case intMinuteIndex
                     	Case 0
	                       strTime = intIndex & ":00 am"
                     	Case 1
	                       strTime = intIndex & ":15 am"
                     	Case 2
	                       strTime = intIndex & ":30 am"
                     	Case 3
	                       strTime = intIndex & ":45 am"
					 End Select	                       
                     strTimeValue = trim(FormatDateTime(strTime, 4))
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
					 Next
                   Next   


				   If strCheckTime = "12:00" Then	
                %>
                	 <option selected value='12:00'>Noon</option>
                <%
				   Else
                %>
                	 <option value='12:00'>Noon</option>
                <%
				   End If



                   For intIndex = 1 To 11 
                     For intMinuteIndex = 0 To 3 
                     Select Case intMinuteIndex
                     	Case 0
	                       strTime = intIndex & ":00 pm"
                     	Case 1
	                       strTime = intIndex & ":15 pm"
                     	Case 2
	                       strTime = intIndex & ":30 pm"
                     	Case 3
	                       strTime = intIndex & ":45 pm"
					 End Select	                       
                     strTimeValue = trim(FormatDateTime(strTime, 4))
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                     
   				   	 If intIndex <> 8 Then	
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
					 End If

				     Next
                   Next   
                %>
                    </select></td>
                  	<td class="style2"></td>
                  </tr>
                  <tr>
                  	<td class="style2">&nbsp;</td>
                  	<td class="style2">&nbsp;</td>
                  	<td class="style2">&nbsp;</td>
                  	<td class="style2">&nbsp;</td>
                  </tr>
                  <tr>
                  	<td>&nbsp;</td>
                  	<td>&nbsp;</td>
                  	<td>&nbsp;</td>
                  	<td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="174" class="style2">
                    &nbsp;
                    Days:</td>
                    <td width="162" class="style2">
                    <input name="days" type="radio" style="width: 20px" value="week" checked="checked">Weekdays:</td>
                    <td width="395">
                    <table width="381" border="0" cellpadding="2" cellspacing="0" >
                      <tr valign="bottom">
                        <td width="20" class="style2">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "1") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="1" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="1"> 
		                <% End If %>
                        </td>
                        <td width="29" class="style2">
                        Sun
                        </td>
                        <td width="20" class="style2">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "2") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="2" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="2"> 
		                <% End If %>
                        </td>
                        <td width="25" class="style2">
                        Mon
                        </td>
                        <td width="20" class="style2">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "3") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="3" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="3"> 
		                <% End If %>
                        </td>
                        <td width="31" class="style2">
                        Tue </td>
                        <td width="20" class="style2">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "4") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="4" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="4"> 
		                <% End If %>
                        </td>
                        <td width="25" class="style2">
                        Wed
                        </td>
                        <td width="20" class="style2">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "5") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="5" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="5"> 
		                <% End If %>
                        </td>
                        <td width="19" class="style2">
                        Thu
                        </td>
                        <td width="20" class="style2">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "6") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="6" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="6"> 
		                <% End If %>
                       </td>
                        <td width="20" class="style2">
                        Fri</td>
                        <td width="20" class="style2">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "7") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="7" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="7"> 
		                <% End If %>
                       </td>
                        <td width="20" class="style2">
                        Sat</td>
                      </tr>
                    </table>
                    </td>
                    <td class="style2"></td>
                  </tr>

<%
'	If err.number <> 0 Then
'	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
'	   response.write "<b>Error occured while creating table<br>"
'	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
'	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
'	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
'	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
'	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
'	End If

%>		
		
	              <tr>
                    <td width="174" class="style2">
                    &nbsp;</td>
                    <td width="162" class="style2">
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    or</td>
                    <td width="395" class="style2">
                    &nbsp;</td>
                    <td class="style2">&nbsp;</td>
                  </tr>

                  <tr>
                    <td width="174" class="style2">
                    &nbsp;</td>
                    <td width="162" class="style2">
                    <input name="days" type="radio" value="fixed">Fixed:</td>
                    <td width="395" class="style2">
                    <input name="Text1" type="text"> mm/dd/yyyy</td>
                    <td class="style2">&nbsp;</td>
                  </tr>

                  <tr>
                    <td width="174" class="style2">
                    &nbsp;</td>
                    <td width="162" class="style2">
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    or</td>
                    <td width="395" class="style2">
                    &nbsp;</td>
                    <td class="style2">&nbsp;</td>
                  </tr>

                  <tr>
                    <td width="174" class="style2">
                    &nbsp;</td>
                    <td width="162" class="style2">
                    <input name="days" type="radio" value="monthly">Monthly:</td>
                    <td width="395" class="style2">
                    The <select name="Select1">
					<option selected="">01</option>
					</select> day of the month</td>
                    <td class="style2">&nbsp;</td>
                  </tr>

                  <tr>
                    <td width="174" class="style2">
                    &nbsp;</td>
                    <td width="162" class="style2">
                    &nbsp;</td>
                    <td width="395" class="style2">
                    &nbsp;</td>
                    <td class="style2">&nbsp;</td>
                  </tr>

	</table>
	<p class="style1"><br><input type="submit" value="  Save  " name="submit">
  <p align="center">&nbsp;</p>
	<input type="hidden" name="schedule_type" value="<%=intScheduleType %>">
	<input type="hidden" name="schedule_id" value="<%=intScheduleID %>">
</form>
<p align="center">&nbsp;</p>
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
			account. <%=intScheduleID %></td>
		</tr>
	</table>
</div>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>