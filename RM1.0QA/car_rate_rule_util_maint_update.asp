<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<%  Response.Expires = -1
    Response.cachecontrol="private" 
    Response.AddHeader "pragma", "no-cache"
    
    'On Error Resume Next

    Server.ScriptTimeout = 30
    
    strURL = "car_rate_rule_util_maint.asp?profile_id=" & Request("profile_id") 

	strConn = Session("pro_con")

	Rem Get the data sources
	Set adoCmd1 = CreateObject("ADODB.Command")
	Set adoCmd2 = CreateObject("ADODB.Command")

	adoCmd1.ActiveConnection =  strConn
	adoCmd1.CommandText = "car_rate_rule_min_utilization_update"
	adoCmd1.CommandType = adCmdStoredProc

	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@rate_rule_id", 3, 1, 0)
	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@util_min_0",   2, 1, 0)
	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@util_min_1",   2, 1, 0)
	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@util_min_2",   2, 1, 0)
	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@util_min_3",   2, 1, 0)
	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@util_min_4",   2, 1, 0)
	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@util_min_5",   2, 1, 0)
	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@util_min_6",   2, 1, 0)
	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@util_min_7",   2, 1, 0)

	adoCmd2.ActiveConnection =  strConn
	adoCmd2.CommandText = "car_rate_rule_max_utilization_update"
	adoCmd2.CommandType = adCmdStoredProc

	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@rate_rule_id", 3, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@util_max_0",   2, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@util_max_1",   2, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@util_max_2",   2, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@util_max_3",   2, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@util_max_4",   2, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@util_max_5",   2, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@util_max_6",   2, 1, 0)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@util_max_7",   2, 1, 0)


	
	For Each Whatever In Request.Form
       
		If Left(Whatever & "          ", 9) = "input_min" Then
		
			strValues = Split(CStr(Request.Form(Whatever)), ",")
		
			adoCmd1.Parameters("@rate_rule_id") = CLng(strValues(0))
			
			For intCount = 0 to 7
				If IsNumeric(strValues(intCount + 1)) Then
					adoCmd1.Parameters("@util_min_" & CStr(intCount)) = CInt(strValues(intCount + 1))
				Else
					adoCmd1.Parameters("@util_min_" & CStr(intCount)) = Null
				End If
			
			Next
			
			adoCmd1.Execute
	
		End If

		If Left(Whatever & "          ", 9) = "input_max" Then
		
			strValues = Split(Request.Form(Whatever) & ",,,,,,,,", ",")
		
			adoCmd2.Parameters("@rate_rule_id") = CLng(strValues(0))
			For intCount = 0 to 7
				If IsNumeric(strValues(intCount + 1)) Then
					adoCmd2.Parameters("@util_max_" & CStr(intCount)) = strValues(intCount + 1)
				Else
					adoCmd2.Parameters("@util_max_" & CStr(intCount)) = Null
				End If
			
			Next

			adoCmd2.Execute

		End If

	Next

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting rule information</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	'Else
	'	Server.Transfer strURL

	End If


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" >
<title>Car Rate Rule Utilization Maint. Update</title>
<style type="text/css">
.style1 {
	text-align: center;
}
.style2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: small;
	text-align: center;
}
</style>
</head>
<body>
<p>&nbsp;</p>
<p>&nbsp;</p>
<%	If err.number = 0 Then  %>
	<p class="style2">Your rules have been updated successfully</p>
<% 	Else		
		For Each Whatever In Request.Form
			Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"
		
		Next
	End If
%>
<form action="<%=strURL %>" method="post" name="return" >
<div class="style1">
	<input type="submit" name="btn_return" value="  Return  " >
</div>
</form>			
</body>
</html>
