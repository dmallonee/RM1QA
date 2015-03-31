<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 30 '180

	On Error Resume Next

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	'intRuleId = Request.QueryString("rateruleid")	
	Dim strDowList
	
	strDowList = Request.Form("dow_list")
	
	strDowList = Replace(strDowList, " ", "")
		
	strConn = Session("pro_con")
	
	If IsNumeric(Request.Form("action")) Then
		intAction = CInt(Request.Form("action"))
	Else
		intAction = 0
	End If


	If intAction > 0 Then
	
		Rem Get the data sources
		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rate_rule_maint"
		adoCmd.CommandType = 4
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_rule_id", 3, 1, 0, Request("rate_rule_id"))
		Select Case intAction
		
			Case 1 'Delete
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_status", 200, 1, 1, "X")
			
			Case 3 'Enable
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_status", 200, 1, 1, "E")
			
			Case 4 'Disable
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_status", 200, 1, 1, "D")
	
		End Select		
	
	Else
	
		Rem Get the data sources
		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_rate_rule_insert"
		adoCmd.CommandType = 4
	
		If (IsNumeric(Request("rate_rule_id"))) And (Request.Form("copy") = "") Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_rule_id", 3, 1, 0, Request("rate_rule_id"))
			intRuleIDAttempt = True
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_rule_id", 3, 1, 0, Null)
			intRuleIDAttempt = False
		End If
		
		adoCmd.Parameters.Append adoCmd.CreateParameter("@alert_desc", 200, 1, 255, Request("alert_desc"))
		
		If IsDate(Request("begin_dt")) Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@begin_dt", 135, 1, 0, Request("begin_dt"))
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@begin_dt", 135, 1, 0, Now)
		End If
		
		If IsDate(Request("end_dt")) Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@end_dt", 135, 1, 0, Request("end_dt"))
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@end_dt", 135, 1, 0, Null)
		End If
	
		If IsDate(Request("first_pickup_dt")) Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@first_pickup_dt", 135, 1, 0, Request("first_pickup_dt"))
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@first_pickup_dt", 135, 1, 0, Now)
		End If
	
		If IsDate(Request("last_pickup_dt")) Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@last_pickup_dt", 135, 1, 0, Request("last_pickup_dt"))
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@last_pickup_dt", 135, 1, 0, Null)
		End If
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@client_sys_rate_cd", 200, 1, 50, Request("client_sys_rate_cd"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cd1", 200, 1, 255,  Request("vend_cd1"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cd2", 200, 1, 255,  Request("vend_cd2"))
		
		adoCmd.Parameters.Append adoCmd.CreateParameter("@lor", 3, 1, 0,  1)
		' Prior version actually used LOR
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@lor", 3, 1, 0,  Request("lor"))
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 255,  Request("city_cd"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd1", 200, 1, 255,  Request("car_type_cd1"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd2", 200, 1, 255,  Request("car_type_cd2"))
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@data_source", 200, 1, 3, "EXP")
		' Prior version actually used data_source
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@data_source", 200, 1, 3,  Request("data_source"))
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@situation_cd", 3, 1, 0,  Request("situation_cd"))
		
		If IsNumeric(Request("situation_amt")) Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@situation_amt", 6, 1, 0, Request("situation_amt"))
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@situation_amt", 6, 1, 0, Null)
		End If
		
		'Rem Phased out, the rules situation and response now contain this information
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@is_dollar", 11, 1, 0, 1)
		If Request("is_dollar") = 1 Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@is_dollar", 11, 1, 0, 1)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@is_dollar", 11, 1, 0, 0)
		End If
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@quantity_period_cd", 3, 1,  0, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@event_count",        2, 1,  0, Null)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@event_time_count",   2, 1,  0, Null)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@event_time_type",  200, 1, 20, Null)
	
		' Prior version used these settings
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@quantity_period_cd", 3, 1, 0,  Request("quantity_period_cd"))
		
		'If IsNumeric(Request("event_count")) Then
		'	adoCmd.Parameters.Append adoCmd.CreateParameter("@event_count", 2, 1, 0, Request("event_count"))
		'Else
		'	adoCmd.Parameters.Append adoCmd.CreateParameter("@event_count", 2, 1, 0, Null)
		'End If
		
		'If IsNumeric(Request("event_time_count")) Then
		'	adoCmd.Parameters.Append adoCmd.CreateParameter("@event_time_count", 2, 1, 0, Request("event_time_count"))
		'Else
		'	adoCmd.Parameters.Append adoCmd.CreateParameter("@event_time_count", 2, 1, 0, Null)
		'End If
		
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@event_time_type", 200, 1, 20,  Request("event_time_type"))
		
		adoCmd.Parameters.Append adoCmd.CreateParameter("@response_cd", 3, 1, 0,  Request("response_cd"))
	
		If IsNumeric(Request("response_amt")) Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@response_amt", 6, 1, 0, Request("response_amt"))
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@response_amt", 6, 1, 0, Null)
		End If
	
		Rem Phased out, the rules situation and response now contain this information
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@is_response_dollar", 11, 1, 0,  Request("is_response_dollar"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@is_response_dollar", 11, 1, 0, 1)
		
		If (Request("search_profile") = 1) And (Request("profile_id") <> "") Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@search_profile", 16, 1, 0, 1)
			adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 200, 1, 4096,  Request("profile_id"))
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@search_profile", 16, 1, 0, 2)
			adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 200, 1, 4096,  "0")
		End If
		
		
		Rem Remember, 99999 is code that we are placing a profile id in the max field.
		If Request("maxmin_profile_id") > 0 Then
				adoCmd.Parameters.Append adoCmd.CreateParameter("@range_max", 6, 1, 0, Request("maxmin_profile_id"))
				adoCmd.Parameters.Append adoCmd.CreateParameter("@range_min", 6, 1, 0, 99999)
		
		Else
			If IsNumeric(Request("rate_maximum")) Then
				adoCmd.Parameters.Append adoCmd.CreateParameter("@range_max", 6, 1, 0, Request("rate_maximum"))
			Else
				adoCmd.Parameters.Append adoCmd.CreateParameter("@range_max", 6, 1, 0, Null)
			End If
		
			If IsNumeric(Request("rate_minimum")) Then
				adoCmd.Parameters.Append adoCmd.CreateParameter("@range_min", 6, 1, 0, Request("rate_minimum"))
			Else
				adoCmd.Parameters.Append adoCmd.CreateParameter("@range_min", 6, 1, 0, Null)
			End If
	
		End If
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_status", 200, 1, 1,  Request("rule_status"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0,  strUserId )
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_rule_type_cd", 3, 1, 0, 1)
		
		Rem a -1 value indicates this rule has it's on success and failure rules stored in the external tables
		Rem not in the field here.
		adoCmd.Parameters.Append adoCmd.CreateParameter("@on_success_id", 3, 1, 0, -1)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@on_failure_id", 3, 1, 0, -1)
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@on_success_id", 3, 1, 0, Request("on_success_id"))
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@on_failure_id", 3, 1, 0, Request("on_failure_id"))
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@utilization_profile_id", 3, 1, 0, Request("utilization_profile_id"))
		
		Dim intUtilCount
		Dim strUtil
		Dim strUtilParam
			
		For intUtilCount = 0 To 7
		
			strUtil = "util_max_" & CStr(intUtilCount)
			strUtilParam = "@util_max_" & CStr(intUtilCount)
			
			If Trim(Request.Form(strUtil)) = "" Then
				adoCmd.Parameters.Append adoCmd.CreateParameter(strUtilParam, 2, 1, 0, Null)
			Else
				adoCmd.Parameters.Append adoCmd.CreateParameter(strUtilParam, 2, 1, 0, Request.Form(strUtil))
			
			End If
	
		Next
		
	
		For intUtilCount = 0 To 7
		
			strUtil = "util_min_" & CStr(intUtilCount)
			strUtilParam = "@util_min_" & CStr(intUtilCount)
			
			If Trim(Request.Form(strUtil)) = "" Then
				adoCmd.Parameters.Append adoCmd.CreateParameter(strUtilParam, 2, 1, 0, Null)
			Else
				adoCmd.Parameters.Append adoCmd.CreateParameter(strUtilParam, 2, 1, 0, Request.Form(strUtil))
			
			End If
	
		Next
	
		If Request("ignore_closed") = "True" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@ignore_closed", 11, 1, 0, 1)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@ignore_closed", 11, 1, 0, 0)
		End If
	
		If Request("rolling_date") = "True" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rolling_date", 11, 1, 0, 1)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rolling_date", 11, 1, 0, 0)
		End If
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@totalprice",       11, 1, 0, 0)
	
		If Request("automatic") = "True" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@automatic",        11, 1, 0, 1)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@automatic",        11, 1, 0, 0)
		End If
		
		adoCmd.Parameters.Append adoCmd.CreateParameter("@dow_list",        200, 1, 13, strDowList)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@comparison_rate",   3, 1, 0, Request.Form("comparison_rate"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rt_amt_tolerance",  6, 1, 0, CCur(Request.Form("rt_amt_tolerance")))
		
		If Request.Form("extra_day_rt") = "" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_day_rt", 6, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_day_rt", 6, 1, 0, Request.Form("extra_day_rt"))
		End If	
	
		If Request.Form("extra_day_miles") = "" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_day_miles", 2, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_day_miles", 2, 1, 0, Request.Form("extra_day_miles"))
		End If	
	
		If Request.Form("extra_day_rt_per_mile") = "" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_day_rt_per_mile", 6, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_day_rt_per_mile", 6, 1, 0, Request.Form("extra_day_rt_per_mile"))
		End If	
	
		If Request.Form("extra_hr_rt") = "" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_hr_rt", 6, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_hr_rt", 6, 1, 0, Request.Form("extra_hr_rt"))
		End If
	
		If Request.Form("extra_hr_miles") = "" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_hr_miles", 2, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_hr_miles", 2, 1, 0, Request.Form("extra_hr_miles"))
		End If
	
		If Request.Form("extra_hr_rt_per_mile") = "" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_hr_rt_per_mile", 6, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@extra_hr_rt_per_mile", 6, 1, 0, Request.Form("extra_hr_rt_per_mile"))
		End If
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_id1",   3, 1, 0, Request.Form("rule_post_action_id1"))
		If Trim(Request.Form("rule_post_action_amt1")) = "" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_amt1", 6, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_amt1", 6, 1, 0, Request.Form("rule_post_action_amt1"))
		End If
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_id2",   3, 1, 0, Request.Form("rule_post_action_id2"))
		If Trim(Request.Form("rule_post_action_amt2")) = "" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_amt2", 6, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_amt2", 6, 1, 0, Request.Form("rule_post_action_amt2"))
		End If
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_id3",   3, 1, 0, Request.Form("rule_post_action_id3"))
		If Trim(Request.Form("rule_post_action_amt3")) = "" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_amt3", 6, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_amt3", 6, 1, 0, Request.Form("rule_post_action_amt3"))
		End If
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_id4",   3, 1, 0, Request.Form("rule_post_action_id4"))
		If Trim(Request.Form("rule_post_action_amt4")) = "" Then
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_amt4", 6, 1, 0, Null)
		Else
			adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_post_action_amt4", 6, 1, 0, Request.Form("rule_post_action_amt4"))
		End If
		
		adoCmd.Parameters.Append adoCmd.CreateParameter("@on_success_id_selected", 200, 1, 4096,  Request("on_success_id_selected"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@on_failure_id_selected", 200, 1, 4096,  Request("on_failure_id_selected"))

	End If

	If Request("debug") <> "true" Then
		Call adoCmd.Execute(,,adExecuteNoRecords)
	
		Rem Comment this out for the new ASPX version
		If err.number = 0 Then
			'Server.Transfer "alerts_rate_management_car_grid.asp"
			Response.Redirect "alerts_rate_management_car_grid.asp"
	
		End If

	End If
	


%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>car_rate_rule_insert</title>
<style type="text/css">
.style1 {
	font-family: Arial, Helvetica, sans-serif;
}
</style>
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
</head>
<body topmargin="0">
<% If (err.number = 0) And (Request("debug") <> "true") Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_middle.jpg">
    <img src="images/top_left.jpg" width="423" height="91" alt=""></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif"><img src="images/top_right.jpg" width="365" height="91" alt=""></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif"></td>
  </tr>
</table>
<p align="right" class="smalltext"><font class="copyright">Copyright (c) 2001-<%=Year(Now)%>, Rate-Highway, Inc. All Rights Reserved.</font></p>
<p>
<p align="center" class="style1"><font size="5" color="#384F5B">Thank you.</font></p>
<p align="center"><font size="5" color="#384F5B"><span class="style1">Your rule has been saved.<br>
</span> </font></p>
<form method="POST" action="return false">
  <p align="center"><input type="button" value=" Close " name="close"  onClick="javascript:window.close();">
</p>
</form>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>


<% Else %>


<b>
<font face="Tahoma">We are sorry, an error has occurred.</font></b><p>
<font face="Tahoma">Please print and fax this page to Rate-Highway Support at (888) 551-0029</font></p>
<p>
<font face="Tahoma"><a href="alerts_rate_management_car.asp">Click here</a> to return to 
               <a href="alerts_rate_management_car.asp">Return to Rate Management</a></font></p>
<p><font size="2" face="Courier New">DEBUG INFORMATION<br>
<hr></font></p>
<p>
<font size="2" face="Courier New">
			   <%		
					If err.number <> 0 Then
					   response.write "<font size='2' face='Courier New'>"
					   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
					   response.write parm_msg & "</b><br>"
					   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
					   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
					   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
					   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
					   response.write pad & "Error Source= <b>" & err.source & "</b><br>"

						i = 0
						For Each error_item in adoCmd.ActiveConnection.Errors
							 response.write pad & adoCmd.ActiveConnection.Errors(i).Description &"<br>"
							 response.write pad & adoCmd.ActiveConnection.Errors(i).NativeError &"<br>"
							 i = i +1
						Next

						Set adoCmd = Nothing
	
						


					   response.write "<hr>"

					End If
					
					
					For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"
       
					Next
					Response.write "Copy = " & Request.Form("copy") & "<br>"					
					response.write intRuleIDAttempt & " <br>"
					response.write adoCmd.Parameters("@rate_rule_id").value

				%>
				
	  		   <br>
				</font>
               </p>

<% End If %>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
	