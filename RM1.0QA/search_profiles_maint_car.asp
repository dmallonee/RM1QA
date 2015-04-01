<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 

<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"
   
   On Error Resume Next

   Server.ScriptTimeout = 0
   
   strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
   strConn = Session("pro_con")   
   
   Select Case Request.Form("maint_action")
   

   	Case "enable"
   	
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_profile_enable"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_profile_id", 3, 1, 0, Request.Form("profile_id"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_status", 200, 1, 1, "E")
	
		adoCmd.Execute
	
		Set adoCmd = Nothing
		
		'Server.Transfer "search_profiles_car.asp"	
	   	
   	Case "disable"
   	
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_profile_enable"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_profile_id", 3, 1, 0, Request.Form("profile_id"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_status", 200, 1, 1, "D")
	
		adoCmd.Execute
	
		Set adoCmd = Nothing
		
		'Server.Transfer "search_profiles_car.asp"	
   	
   	
   	Case "delete"
   	
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_profile_enable"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_profile_id", 3, 1, 0, Request.Form("profile_id"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_status", 200, 1, 1, "X")
	
		adoCmd.Execute
	
		Set adoCmd = Nothing
		
		'Server.Transfer "search_profiles_car.asp"	
   	
	Case "copy_new_user"

		strConn = Session("pro_con")
	
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_profile_copy"
		adoCmd.CommandType = 4


		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Request.Form("profile_id"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255, Null)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0, Request.Form("users"))

		adoCmd.Execute   		
   	
   		'Server.Transfer "search_profiles_car.asp"
	
	
	

   	Case "copy"
		strConn = Session("pro_con")
	
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_profile_copy"
		adoCmd.CommandType = 4


		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Request.Form("profile_id"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255, Request.Form("new_name"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0, Null)

		adoCmd.Execute   		
   	
   		'Server.Transfer "search_profiles_car.asp"
   	
   	

   	Case "rename"
		strConn = Session("pro_con")
	
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_profile_rename"
		adoCmd.CommandType = 4


		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Request.Form("profile_id"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255, Request.Form("new_name2"))

		adoCmd.Execute   		
   	
   		'Server.Transfer "search_profiles_car.asp"
   	
   	
   End Select


   
   
	If err.number = 0 Then
		Server.Transfer "search_profiles_car.asp"	
	Else
			
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
			
		Session("error_msg") = "An error was encountered while request your search. Please contact Rate-Highway support"
		'Server.Transfer "search_criteria_car.asp"
	End If

   
   

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>search_profiles_maint_car.asp</title>
</head>
<body>
		<%	For Each Whatever In Request.Form
				Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"
       
			Next
		%>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
