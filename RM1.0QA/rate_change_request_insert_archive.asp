<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<% Response.Expires = -1
   Response.cachecontrol = "private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 0
   
   On Error Resume Next
 
	Dim strConn	
	Dim adoCmd	
	Dim adoRS

	strConn = Session("pro_con")
	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_notify_support_rate_changes"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@report_request_id", 3, 1, 0, Request.Form("reportrequestid"))

	Call adoCmd.Execute(,,adExecuteNoRecords)
	
	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write "Error notifying support</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If

	

	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_rate_change_queue_insert"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_change_id", 3, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_amt", 6, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request("reportrequestid"))
	
	Dim strArray
	Dim intArrayHolder
	Dim strRateArray
	
	strArray = Split(Request.Form("car_rate_rule_change_id"), ",", -1, 1)
	strRateArray = Split(Request.Form("new_rate_amt"), ",", -1, 1)
	
	For intArrayHolder = LBound(strArray) To UBound(strArray)

		adoCmd.Parameters("@car_rate_rule_change_id").Value = strArray(intArrayHolder)

		If intArrayHolder > 0 Then		
			If InStr(1, strRateArray(intArrayHolder), "@") Then
				adoCmd.Parameters("@rate_amt").Value = Replace(strRateArray(intArrayHolder), "@", "")
			Else
				adoCmd.Parameters("@rate_amt").Value = Null
			End If 
		Else
			adoCmd.Parameters("@rate_amt").Value = Null
		End If 
		
		'Call adoCmd.Execute(,,adExecuteNoRecords)

	
	Next

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write "Error inserting updates</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If


If False Then
	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_rate_change_insert"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_change_id", 3, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_amt", 6, 1, 0, Null)
	
	Dim strArray
	Dim intArrayHolder
	Dim strRateArray
	
	strArray = Split(Request.Form("car_rate_rule_change_id"), ",", -1, 1)
	strRateArray = Split(Request.Form("new_rate_amt"), ",", -1, 1)
	
	For intArrayHolder = LBound(strArray) To UBound(strArray)

		adoCmd.Parameters("@car_rate_rule_change_id").Value = strArray(intArrayHolder)

		If intArrayHolder > 0 Then		
			If InStr(1, strRateArray(intArrayHolder), "@") Then
				adoCmd.Parameters("@rate_amt").Value = Replace(strRateArray(intArrayHolder), "@", "")
			Else
				adoCmd.Parameters("@rate_amt").Value = Null
			End If 
		Else
			adoCmd.Parameters("@rate_amt").Value = Null
		End If 
		
		Call adoCmd.Execute(,,adExecuteNoRecords)

	
	Next

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write "Error inserting updates</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If


	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "car_shop_request_org_detail"
	adoCmd.CommandType = 4
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request("reportrequestid"))

	Set adoRS = adoCmd.Execute
		
	Select Case (adoRS.Fields("parent_id").Value)
	
		Case 6 'Payless only
		
		'If (adoRS.Fields("parent_id").Value = 6) Then
		Set adoCmd = CreateObject("ADODB.Command")
		adoCmd.ActiveConnection = strConn
		adoCmd.CommandText = "car_rate_change_queue_payless_insert_condensed"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id1", 3, 1, 0, Request.Form("reportrequestid"))
		Call adoCmd.Execute(,,adExecuteNoRecords)
		
		If err.number <> 0 Then
		   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
		   response.write "<b>VBScript Errors Occured!<br>"
		   response.write "Error in custom section</b><br>"
		   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
		   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
		   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
		   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
		   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
		End If
		
		
		
		'End If
		
	End Select
	
	Set adoRS = Nothing

	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "car_notify_user_rate_change_receipt"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request.Form("reportrequestid"))
	Call adoCmd.Execute(,,adExecuteNoRecords)

End If

	Set adoCmd = Nothing
	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	Else
		Server.Transfer "rate_change_complete.asp"
	
	End If


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>search_request_insert</title>
</head>
<body>
We are sorry, an error has occurred.<p>Please print and fax this page to 
Rate-Highway Support at (888) 551-0029</p>
<p><font size="2">Debug information</font></p>
<p><br>


			   <%	For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"

       
					Next
				%>

</font></p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>

