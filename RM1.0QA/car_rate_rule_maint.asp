<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180

	'On Error Resume Next

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	intArray = Split(Request.Form("rate_rule_id"), ",")
	intCount = UBound(intArray)
	
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_rate_rule_maint"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_rule_id", 3, 1, 0)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@rule_status", 200, 1, 1)


	While intCount >= 0
	
		If intArray(intCount) <> "" Then
			adoCmd.Parameters("@rate_rule_id").Value = intArray(intCount)
			Select Case Request.Form("action")
			
				Case 1
					adoCmd.Parameters("@rule_status").Value = "X"
			
				Case 3
					adoCmd.Parameters("@rule_status").Value = "E"
	
				Case 4
					adoCmd.Parameters("@rule_status").Value = "D"
	
				Case 5
					Rem S = Sandbox
					adoCmd.Parameters("@rule_status").Value = "S"
	
				Case Else
					adoCmd.Parameters("@rule_status").Value = "E"
	
			End Select

			Call adoCmd.Execute
			
		End If
		intCount = intCount - 1
	
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


	i = 0
	For Each error_item in adoCmd.ActiveConnection.Errors
	 response.write adoCmd.ActiveConnection.Errors(i).Description &"<br>"
	 response.write adoCmd.ActiveConnection.Errors(i).NativeError &"<br>"
	 i = i +1
	Next

	Response.Redirect "alerts_rate_management_car.asp"
	
	

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>rate_rule_mainta</title>
</head>
<body>

			   <%	For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"

       
					Next
					
					Response.Write UBound(intArray)

				%>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
	
