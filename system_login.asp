Ente<%@ Language=VBScript %>
<!--
Revisions
When     Who What
========== ==== ==========================================================
10/22/2014 DLM  Modified to define connection string based on database variables
                Purpose is to eliminate any requirement that there be multiple websites.
-->
<!-- #INCLUDE FILE="include/adovbs.asp" --> 
<%
    Response.Expires = -1
	Response.Expiresabsolute = Now() - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"

	Dim strLoginCode
	Dim strPassword 
	Dim strRememberMe 
    Dim strURL

	Dim adoCmd
    Dim adoRS
	Dim objItem 
	Dim intParentID

	strLoginCode  = Request("email_address")
	strPassword   = Request("password")
	strRememberMe = Request("Remember")
	strURL        = Request("request_url")
	
	Session("user_name") = "Not Logged In"

	Set adoCmd = Server.CreateObject("ADODB.Command")
   	With adoCmd
        .ActiveConnection = Session("pro_con_login")
		.CommandText = "user_login"
		.CommandType = adCmdStoredProc
	End with

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  adVarChar, adParamInput, 50, strLoginCode )
    adoCmd.Parameters.Append adoCmd.CreateParameter("@password", adVarChar, adParamInput, 50, strPassword )
	
	Set adoRS = adoCmd.Execute

	If adoRS.EOF = False Then
        'REDIRECT TO SEND DESIGNATED USERS TO THE BETA WEBSITE
		'CURRENTLY DISABLED FOR SUPPORT
        if adoRS.Fields("user_id").Value = 33 and Session("user_level") = "0" Then
        '    Response.Redirect "http://preview.rate-monitor.com/default.aspx?email_address=" & strLoginCode & "&password=" & strPassword
        '   Response.End
        End if
                 
		If adoRS.Fields("user_id").Value > 0 Then

       	    Response.Cookies("rmuserid") = adoRS.Fields("user_id").Value		
		    Response.Cookies("rmusername") = adoRS.Fields("first_name").Value & " " & adoRS.Fields("last_name").Value
		    Response.Cookies("loginCode") = strLoginCode 
		    Response.Cookies("password") = strPassword
		    Response.Cookies("remember") = True
		    Response.Cookies("testing") = adoRS.Fields("testing").Value
		    Response.Cookies("user_id") = adoRS.Fields("user_id").Value
		    Response.Cookies("user_name") = adoRS.Fields("first_name").Value & " " & adoRS.Fields("last_name").Value
		    Response.Cookies("client_userid") = adoRS.Fields("client_userid").Value
		    Response.Cookies("rpt_limit") = adoRS.Fields("rpt_limit").Value
		    Response.Cookies("user_role") = adoRS.Fields("user_role").Value
REM Exception for Fox 1-way
    if strLoginCode = "fx1way" then
    	    Session("pro_con") = "Provider=SQLOLEDB; Network Library=dbmssocn;Password=iLOVEtab@sco!;User ID=rhWeb;Initial Catalog=prod_fx1;Data Source=athena.rate-monitor.com;"
    else								
    	    Session("pro_con") = "Provider=SQLOLEDB; Network Library=dbmssocn;Password=iLOVEtab@sco!;User ID=rhWeb;Initial Catalog=" & adoRS.Fields("dbname").Value & ";Data Source=" & adoRS.Fields("dbserver").Value & ".rate-monitor.com;"
    end if

		    Session("testing") = adoRS.Fields("testing").Value
		    Session("user_name") = adoRS.Fields("first_name").Value & " " & adoRS.Fields("last_name").Value
		    Session("user_id") = adoRS.Fields("user_id").Value
		    Session("org_id") = adoRS.Fields("org_id").Value
		    Session("us_date") = adoRS.Fields("us_date").Value
		    Session("us_decimal") = adoRS.Fields("us_decimal").Value
            Session("password") = strPassword
            Session("user_level") = adoRS.Fields("user_level").Value
			Session("user_role") = adoRS.Fields("user_role").Value
            Session("maintenance") = adoRS.Fields("maintenance").Value
            Session("minmax") = adoRS.Fields("minmax_on").Value
            Session("tp") = adoRS.Fields("totalprice_on").Value    
            Session("threshold") = adoRS.Fields("threshold_on").Value
            Session("boomerang") = adoRS.Fields("boomerang_on").Value    
            Session("display_rates") = adoRS.Fields("display_rates").Value
            Session("division_id") = adoRS.Fields("division_id").Value
            Session("site") = Request.ServerVariables("SERVER_NAME")
 
			If adoRS.Fields("enabled").Value = "False" Then
			'    Server.Transfer "login_suspended.asp"
				Response.Redirect "login_suspended.asp"
				Response.End
		    End If
		
		    intParentID = adoRS.Fields("parent_id").Value

		    Set adoRS = adoRS.NextRecordset
		
		    For Each objItem in adoRS.Fields
			    If (strPassword = strSupportPwd ) Then
				    Rem if the user is from support, allow them access to all areas
				    If objItem.Name = "user_id" Then
					    Rem all fields are t/f flags except for the user_id so we have to make an exception
					    Response.Cookies(objItem.Name) = objItem.Value
				    Else
					    Response.Cookies(objItem.Name) = True
				    End If
				    Session("support") = 1
			    Else
                REM    if objItem.Name <> "user_guid" then
				        Response.Cookies(objItem.Name) = objItem.Value
                REM    end if
			    End If
 		    Next

		    If strURL <> "" Then
 '   		    Server.Transfer strURL
				Response.Redirect strURL
				Response.End
 
		    Else
                if	Session("user_level") > 0 and Session("maintenance") = "True" then
'                    Server.Transfer "maintenance.asp"
					Response.Redirect "maintenance.asp"
					Response.End
	            else 
 '   				Server.Transfer "login_welcome.asp"
					Response.Redirect "login_welcome.asp"
					Response.End
                end if
		    End If
        End if

    Else
'		Server.Transfer "login_error.asp"
		Response.Redirect "login_error.asp"
		Response.End
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

r file contents here
