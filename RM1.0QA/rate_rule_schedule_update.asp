<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"

    Server.ScriptTimeout = 360

	On Error Resume Next

	strUserId = Request.Cookies("rate-monitor.com")("user_id")


	Rem Dump all the rows currently held by this schedule_id, then re-insert them from the webpage

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim intScheduleID
	Dim strType
	Dim strFiller
	Dim strValues
	Dim strCount
	Dim strUserId

	intScheduleID = Request.Form("schedule_id")
	intScheduleType = Request.Form("schedule_type")
	strUserId = Request.Cookies("rate-monitor.com")("user_id")

	strType = Request("type")
	strFiller = ", , , , , , , , , "
	
	If IsNumeric(intScheduleID) = False Then
		intScheduleID = 0
	End If
	
	Dim intDaysOutGrp1
	Dim intDaysOutGrp2
	Dim intDaysOutGrp3
	Dim intDaysOutGrp4
	Dim intDaysOutGrp5
	Dim intDaysOutGrp6
	Dim intDaysOutGrp7
	Dim intDaysOutGrp8
	Dim intDaysOutGrp9
	
	intDaysOutGrp1 = 0 'Request.Form("days_out_grp1")
	intDaysOutGrp2 = 1 'Request.Form("days_out_grp2")
	intDaysOutGrp3 = 2 'Request.Form("days_out_grp3")
	intDaysOutGrp4 = 3 'Request.Form("days_out_grp4")
	intDaysOutGrp5 = 4 'Request.Form("days_out_grp5")
	intDaysOutGrp6 = 5 'Request.Form("days_out_grp6")
	intDaysOutGrp7 = 6 'Request.Form("days_out_grp7")
	intDaysOutGrp8 = 7 'Request.Form("days_out_grp8")
	
	'Not used right now
	'intDaysOutGrp9 = Request.Form("days_out_grp9")
	
		strConn = Session("pro_con")
	
	  	Set adoRS = CreateObject("ADODB.Recordset")

		Rem First, update the schedule header, just in case the user changed the description/name
	  	Set adoCmd = CreateObject("ADODB.Command")
		
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_rate_rule_schedule_header_insert"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("RETURN_VALUE",                   3, adParamReturnValue)
		If intScheduleID = 0 Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_schedule_id",     3, 1,  0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_schedule_id",     3, 1,  0, intScheduleID)
		End If
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_schedule_desc", 200, 1, 50, Request.Form("name"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",                       3, 1,  0, strUserId )
		adoCmd.Parameters.Append adoCmd.CreateParameter("@schedule_type_id",              3, 1,  0, intScheduleType )
		
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
		End If

		intScheduleID = adoCmd.Parameters("RETURN_VALUE").Value
	
		Rem Second, purge out all the detail rows so we can write in the new ones
	  	Set adoCmd = CreateObject("ADODB.Command")
		
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_rate_rule_schedule_detail_1_delete"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_schedule_id", 3, 1, 0, intScheduleID)

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
		End If

	  	Set adoCmd = CreateObject("ADODB.Command")
		
		Rem Third, loop around and insert all the new items
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_rate_rule_schedule_detail_1_insert"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_schedule_id", 3, 1, 0, intScheduleID)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@days_out_grp",             17, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@util_min",                  2, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@response_amt",              6, 1, 0)

		Dim strCell
		Dim intCounter
		
		intCounter = 1
		strCell = Request.Form("cell" & intCounter) & ""
		
		While (Len(strCell) > 0)
		

			If (strCell  <> strFiller) And (Len(strCell) > 0) Then
				strValues = Split(strCell, ",")
				For intCount = LBound(strValues) + 1 to UBound(strValues)
					If Len(strValues(intCount) & "") > 0 Then

						If IsNumeric(strValues(intCount)) Then

						adoCmd.Parameters("@util_min").Value = strValues(0)
						adoCmd.Parameters("@response_amt").Value = strValues(intCount)

						Select Case intCount
							Case 1 
								'Response.Write Request.Form("days_out_grp1") & "/" & strValues(0) & "/" & strValues(intCount) & "/0<br>"
								adoCmd.Parameters("@days_out_grp").Value = intDaysOutGrp1 
								adoCmd.Execute
							Case 2 
								'Response.Write Request.Form("days_out_grp2") & "/" & strValues(0) & "/" & strValues(intCount) & "/1<br>"
								adoCmd.Parameters("@days_out_grp").Value = intDaysOutGrp2
								adoCmd.Execute
							Case 3 
								'Response.Write Request.Form("days_out_grp3") & "/" & strValues(0) & "/" & strValues(intCount) & "/2<br>"
								adoCmd.Parameters("@days_out_grp").Value = intDaysOutGrp3 
								adoCmd.Execute
							Case 4 
								'Response.Write intDaysOutGrp4 & "/" & strValues(0) & "/" & strValues(intCount) & "/3<br>"
								adoCmd.Parameters("@days_out_grp").Value = intDaysOutGrp4 
								adoCmd.Execute
							Case 5 
								'Response.Write intDaysOutGrp5 & "/" & intCounter & "/" & strValues(intCount) & "/4<br>"
								adoCmd.Parameters("@days_out_grp").Value = intDaysOutGrp5 
								adoCmd.Execute
							Case 6 
								'Response.Write intDaysOutGrp6 & "/" & intCounter & "/" & strValues(intCount) & "/5<br>"
								adoCmd.Parameters("@days_out_grp").Value = intDaysOutGrp6 
								adoCmd.Execute
							Case 7 
								'Response.Write intDaysOutGrp7 & "/" & intCounter & "/" & strValues(intCount) & "/6<br>"
								adoCmd.Parameters("@days_out_grp").Value = intDaysOutGrp7 
								adoCmd.Execute
							Case 8 
								'Response.Write intDaysOutGrp8 & "/" & intCounter & "/" & strValues(intCount) & "/7<br>"
								adoCmd.Parameters("@days_out_grp").Value = intDaysOutGrp8 
								adoCmd.Execute
'							Case 9
'								Exit For
						End Select								

						End If

						
					End If
	
				Next
			
			End If

			intCounter = intCounter + 1
			
			strCell = Request.Form("cell" & intCounter) & ""

		
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
<font size="4">Your Schedule changes are complete.</font><p align="center">
&nbsp;<form method="POST" action="rate_rule_schedule_update.asp" webbot-action="--WEBBOT-SELF--">
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
</body>

</html>