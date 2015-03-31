<%@ Language=VBScript %>
<!--
Revisions
When     Who What
======== === ==========================================================
-->
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<%
	On Error Resume Next

    Response.Expires = -1
	Response.Expiresabsolute = Now() - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
	'Response.Buffer = True 
 
	Dim strLoginCode
	Dim strPassword 
	Dim strRememberMe 
	Dim strFromForm
	Dim strConn
	Dim strSQL
	Dim adoCmd
	Dim ReturnValue
	Dim Database
	Dim objItem 
	Dim intLOB_id
	Dim strURL
	Dim intParentID
	Dim strSite
	Dim website

    Dim strConn2
    Dim adoCmd2
    Dim adoRS2
    	
	strLoginCode  = Request.Form("email_address") & ""
	strPassword   = Request.Form("password") & ""
	strRememberMe = Request.Form("Remember") & ""
	strFromForm   = Request.Form("FromForm") & ""
	intDatabase   = Request.Form("database") & ""
	strURL        = Request.Form("request_url") & ""
	strSite       = Request.ServerVariables("SERVER_NAME")
	
	Session("user_name") = "Not Logged In"


	Set adoCmd = Server.CreateObject("ADODB.Command")
   	With adoCmd
		.ActiveConnection = Session("pro_con")
REM                .ActiveConnection = Session("pro_con_login")
		.CommandText = "user_login"
		.CommandType = adCmdStoredProc
	End with
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  adVarChar, adParamInput, 50, strLoginCode )
    adoCmd.Parameters.Append adoCmd.CreateParameter("@password", adVarChar, adParamInput, 50, strPassword )
	
	Set adoRS = adoCmd.Execute
	
	strReturnValue = "failure"
	
	If adoRS.State = adStateClosed Then
		Response.Cookies("rate-monitor.com").Domain = "rate-monitor.com"	 		
		Response.Cookies("rate-monitor.com").Path = "/"	 		

 		Response.Cookies("rate-monitor.com") = ""
		Response.Cookies("rmuserid") = ""
        Response.Cookies("rmusername") = ""

		Session("testing") = "false"
		Session("user_name") = ""
		Session("user_id") = 0
		Session("org_id") = 0
		Session("msg") = "result set is closed:" & strLoginCode & ":" & strPassword 

		Server.Transfer "login_error.asp"
		
	ElseIf adoRS.EOF = False Then
        'REDIRECT TO SEND DESIGNATED USERS TO THE BETA WEBSITE
        if adoRS.Fields("user_id").Value = 33 and Session("user_level") = "0" Then
            Response.Redirect "http://preview.rate-monitor.com/default.aspx?email_address=" & strLoginCode & "&password=" & strPassword
            Response.End
        End if
REM     website = adoRS.Fields("website").Value & ".rate-monitor.com"
REM     Session("dbserver") = adoRS.Fields("dbserver").Value
REM     If website <> strSite Then
REM         qs = "http:/" & website & "/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
REM         Response.Redirect qs
REM         Response.End
REM     End If                  
		If adoRS.Fields("user_id").Value > 0 Then
			If (adoRS.Fields("server_name").Value = "HERCULES") And (strSite <> "ehi.rate-monitor.com") Then
				qs = "http://ehi.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
                Response.Redirect qs
                Response.End
			ElseIf (adoRS.Fields("server_name").Value = "CRONOS") And (strSite <> "ehi.rate-monitor.com") Then
				qs = "http://ehi.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
			    Response.Redirect qs
			    Response.End		
			ElseIf (adoRS.Fields("server_name").Value = "THOREHI") And (strSite <> "ehi.rate-monitor.com") Then
				qs = "http://ehi.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
				Response.Redirect qs
			    Response.End	
			ElseIf (adoRS.Fields("server_name").Value = "EHI") And (strSite <> "ehi.rate-monitor.com") Then
				qs = "http://ehi.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
				Response.Redirect qs
			    Response.End	
			'ElseIf (adoRS.Fields("server_name").Value = "THOR") And (strSite <> "www.rate-monitor.com") Then
			'	qs = "http://www.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
 			'    Response.Redirect qs
			'    Response.End	
			ElseIf (adoRS.Fields("server_name").Value = "ATHENA") And (strSite <> "fox.rate-monitor.com") Then
				qs = "http://fox.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
			    Response.Redirect qs
			    Response.End	
			ElseIf (adoRS.Fields("server_name").Value = "RHEA") And (strSite <> "advantage.rate-monitor.com") Then
				qs = "http://advantage.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
			    Response.Redirect qs
			    Response.End	
			ElseIf (adoRS.Fields("server_name").Value = "ADVANTAGE") And (strSite <> "advantage.rate-monitor.com") Then
				qs = "http://advantage.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
			    Response.Redirect qs
			    Response.End	
			ElseIf (adoRS.Fields("server_name").Value = "FOX") And (strSite <> "fox.rate-monitor.com") Then
				qs = "http://fox.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
			    	Response.Redirect qs
			    	Response.End	
            		ElseIf (adoRS.Fields("server_name").Value = "ABG") And (strSite <> "abg.rate-monitor.com") Then
				qs = "http://abg.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
			    	Response.Redirect qs
			    	Response.End		
            		ElseIf (adoRS.Fields("server_name").Value = "PAYLESS") And (strSite <> "payless.rate-monitor.com") Then
				qs = "http://payless.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
			    	Response.Redirect qs
			    	Response.End      	
            		ElseIf (adoRS.Fields("server_name").Value = "WWW") And (strSite <> "www.rate-monitor.com") Then
				qs = "http://www.rate-monitor.com/system_login_transfer.asp?email_address=" & strLoginCode & "&password=" & strPassword
			    	Response.Redirect qs
			    	Response.End		
            		End If
		End If

        strConn2 = Session("pro_con")
REM     strConn2 = Session("pro_con_login")
  		Set adoCmd2 = CreateObject("ADODB.Command")
		adoCmd2.ActiveConnection =  strConn2
		adoCmd2.CommandText = "system_management_variables_select"
		adoCmd2.CommandType = 4
		adoCmd2.Parameters.Append adoCmd.CreateParameter("@prod_server",      200, 1, 20)
		adoCmd2.Parameters.Append adoCmd.CreateParameter("@database_name",      200, 1, 20)
		adoCmd2.Parameters("@prod_server").Value = "ATHENA"
        adoCmd2.Parameters("@database_name").Value = "production"
REM		adoCmd2.Parameters("@prod_server").Value = Session("dbserver") 
        Set adoRS2 = Server.CreateObject("ADODB.Recordset")
		adoRS2.Open adoCmd2, , adOpenStatic, adLockReadOnly
        start = adoRS2("maintenance_start").Value 
        duration = adoRS2("maintenance_duration").Value
        starttime = split(start,":")
        'MAINTENANCE WINDOW DISABLED
        If (Hour(now) = starttime(0)) And (Minute(now) < duration) And (Session("user_level") <> "0") Then
           ' Server.Transfer "maintenance.asp"
           ' Response.End
    	End If

        'If (Hour(now) = 20) And (Session("user_level") <> "0") Then
  		'    Server.Transfer "maintenance.asp"
	    'End If

       	Response.Cookies("rmuserid") = adoRS.Fields("user_id").Value		
		Response.Cookies("rmusername") = adoRS.Fields("first_name").Value & " " & adoRS.Fields("last_name").Value

		Response.Cookies("rate-monitor.com").Domain = strSite	 		
		Response.Cookies("rate-monitor.com").Path = "/"	 		
 		Response.Cookies("rate-monitor.com")("live_session") = "auto"
		Response.Cookies("rate-monitor.com")("loginCode") = strLoginCode 
		Response.Cookies("rate-monitor.com")("password") = strPassword
		Response.Cookies("rate-monitor.com")("remember Me") = True
		Response.Cookies("rate-monitor.com")("testing") = adoRS.Fields("testing").Value
		Response.Cookies("rate-monitor.com")("user_id") = adoRS.Fields("user_id").Value
		Response.Cookies("rate-monitor.com")("user_name") = adoRS.Fields("first_name").Value & " " & adoRS.Fields("last_name").Value
		Response.Cookies("rate-monitor.com")("client_userid") = adoRS.Fields("client_userid").Value
		Response.Cookies("rate-monitor.com").Expires = "July 4, 2016"
		Response.Cookies("rate-monitor.com")("vend_cd") = adoRS.Fields("self_vend_cd").Value			
		Response.Cookies("rate-monitor.com")("rpt_limit") = adoRS.Fields("rpt_limit").Value
		Response.Cookies("rate-monitor.com")("user_role") = adoRS.Fields("user_role").Value
								
		Session("testing") = adoRS.Fields("testing").Value
		Session("user_name") = adoRS.Fields("first_name").Value & " " & adoRS.Fields("last_name").Value
		Session("user_id") = adoRS.Fields("user_id").Value
		Session("org_id") = adoRS.Fields("org_id").Value
		Session("us_date") = adoRS.Fields("us_date").Value
		Session("us_decimal") = adoRS.Fields("us_decimal").Value
        Session("password") = strPassword
        Session("user_level") = adoRS.Fields("found_in").Value
        Session("minmax") = adoRS.Fields("update_minmax_sched").Value
        Session("threshold") = adoRS.Fields("update_threshold_sched").Value
        Session("tp") = adoRS.Fields("update_TP_minmax_sched").Value    
        Session("display_rates") = adoRS.Fields("display_rates").Value
        Session("division_id") = adoRS.Fields("division_id").Value
                    	
		If adoRS.Fields("expiration").Value <= Now Then
			Server.Transfer "suspended.asp"
							
		End If
		
		intLOB_id = adoRS.Fields("lob_id").Value
		intParentID = adoRS.Fields("parent_id").Value

		Set adoRS = adoRS.NextRecordset
		
		For Each objItem in adoRS.Fields
			If (strPassword = strSupportPwd ) Then
				Rem if the user is from support, allow them access to all areas
				If objItem.Name = "user_id" Then
					Rem all fields are t/f flags except for the user_id so we have to make an exception
					Response.Cookies("rate-monitor.com")(objItem.Name) = objItem.Value
				Else
					Response.Cookies("rate-monitor.com")(objItem.Name) = True
				End If
				Session("support") = 1
			Else
            REM    if objItem.Name <> "user_guid" then
				    Response.Cookies("rate-monitor.com")(objItem.Name) = objItem.Value
            REM    end if
			End If
 	
		
		Next
		
		Session("server_name") = strSite 

		If (Request.Cookies("rate-monitor.com")("monthly_amt") = "True") Then
			Response.Redirect "login_welcome.asp"
REM			Response.Redirect "maintenance.asp"
			
		End If

		If strFromForm = "report_login" Then
			If strURL <> "" Then
				Server.Transfer strURL
			Else
				Server.Transfer "system_error.asp"
			End If
		Else					
			Response.Redirect "search_queue_car.asp"
REM			Response.Redirect "maintenance.asp"
		End If
		
	
	Else
		strReturnValue = "empty resultset"
		Server.Transfer "login_error.asp"
	
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
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Administration | Login</title>
</head>
<body>
<!--  
<form action="" method="post" name="login">
	<input name="password" type="hidden" value="<%=strPassword %>">
	<input name="email_address" type="hidden" value="<%=strLoginCode %>">
	<input name="Remember" type="hidden" value="<%=strRememberMe %>">
	<input name="FromForm" type="hidden" value="<%=strFromForm %>">
	<input name="database" type="hidden" value="<%=intDatabase %>">
	<input name="request_url" type="hidden" value="<%=strURL %>">
</form> -->
</body>
</html>