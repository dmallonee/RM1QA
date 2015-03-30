<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<% Response.Expires = -1
   Response.cachecontrol = "private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 0
   
   'On Error Resume Next
 
	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim debug
	Dim strArray
	Dim intArrayHolder
	Dim strRateArray
	Dim strRateString
	Dim strErrorListing
	Dim strIdString 
	
	strErrorListing = ""

	strConn = Session("pro_con")

	Session("reportrequestid") =  Request("reportrequestid")
	Session("security_code") =  Request("security_code")

	
	debug = Request.Form("debug")
	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   strErrorListing = strErrorListing & pad & "<b>Error notifying support</b><br>"
	   strErrorListing = strErrorListing & pad & "Error Number   = #<b>" & err.number & "</b><br>"
	   strErrorListing = strErrorListing & pad & "Error Desc.    = <b>" & err.description & "</b><br>"
	   strErrorListing = strErrorListing & pad & "Help Context   = <b>" & err.HelpContext & "</b><br>"
	   strErrorListing = strErrorListing & pad & "Help File Path = <b>" & err.helpfile & "</b><br>"
	   strErrorListing = strErrorListing & pad & "Error Source   = <b>" & err.source & "</b><br><hr>"
	End If

	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "car_rate_change_queue_insert"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@car_rate_rule_change_id", 3, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@rate_amt",                6, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id",         3, 1, 0, Request("reportrequestid"))

	strRateChangeId = Request.Form("car_rate_rule_change_id")		
	strArray = Split(strRateChangeId, ",", -1, 1)
	
	strRateString = Request.Form("new_rate_amt")
	strRateArray = Split(strRateString, ",", -1, 1)


	If (debug = 1) Then
		Response.Write "DEBUG - " & strRateChangeId  & ", " &  strRateString & "<BR>"
	End If


	If InStr(1, strRateChangeId, ",") = 0 Then
		adoCmd.Parameters("@car_rate_rule_change_id").Value = strRateChangeId 		
		If (Len(strRateString) > 0) Then		
			If InStr(1, strRateString, "@") Then
				adoCmd.Parameters("@rate_amt").Value = Replace(strRateString, "@", "")
			Else
				adoCmd.Parameters("@rate_amt").Value = Null
			End If 
		Else
			adoCmd.Parameters("@rate_amt").Value = Null
		End If 

		Call adoCmd.Execute(,,adExecuteNoRecords)
		
	Else
		
		For intArrayHolder = LBound(strArray) To UBound(strArray)
		
			If (debug = 1) Then
				Response.Write "DEBUG - " & intArrayHolder & " - " & strArray(intArrayHolder) & ", " &  strRateArray(intArrayHolder) & "<BR>"
			End If
	
			adoCmd.Parameters("@car_rate_rule_change_id").Value = strArray(intArrayHolder)
	
			If (intArrayHolder >= 0) And (Len(strRateString) > 0) Then		
				If InStr(1, strRateArray(intArrayHolder), "@") Then
					adoCmd.Parameters("@rate_amt").Value = Replace(strRateArray(intArrayHolder), "@", "")
					If (debug = 1) Then
						Response.Write "override amount = " &  Replace(strRateArray(intArrayHolder), "@", "") & "<br>"
					End If
	
				Else
					adoCmd.Parameters("@rate_amt").Value = Null
					If (debug = 1) Then
						Response.Write "Not at sign in amount = " &  strRateArray(intArrayHolder) & "<br>"
					End If
	
				End If 
			Else
				adoCmd.Parameters("@rate_amt").Value = Null
			End If 
	
	
	'		If (debug = 1) Then
	'			Response.Write adoCmd.CommandText & " " & adoCmd.Parameters("@car_rate_rule_change_id").Value & ", " &  (adoCmd.Parameters("@rate_amt").Value & ", " & adoCmd.Parameters("@shop_request_id").Value & "<BR>"
	'		End If
			
			Call adoCmd.Execute(,,adExecuteNoRecords)
	
		
		Next

	End If

	Set adoCmd = Nothing
	
	If (err.number <> 0) Or (debug = 1) Then
		pad="&nbsp;&nbsp;&nbsp;&nbsp;"
		strErrorListing = strErrorListing & pad & "<b>Error adding rate change to queue</b><br>"
		strErrorListing = strErrorListing & pad & "Error Number   = #<b>" & err.number & "</b><br>"
		strErrorListing = strErrorListing & pad & "Error Desc.    = <b>" & err.description & "</b><br>"
		strErrorListing = strErrorListing & pad & "Help Context   = <b>" & err.HelpContext & "</b><br>"
		strErrorListing = strErrorListing & pad & "Help File Path = <b>" & err.helpfile & "</b><br>"
		strErrorListing = strErrorListing & pad & "Error Source   = <b>" & err.source & "</b><br><hr>"
	Else
		Server.Transfer "rate_change_complete.asp"
	
	End If


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>search_request_insert</title>
<style type="text/css">
<!--
.copyright {
	FONT-SIZE: 0.7em; TEXT-ALIGN: right
}
-->
</style>
</head>
<body style="font-family: Tahoma; font-size: 10pt">
<table width="100%" border="0" cellspacing="0" cellpadding="0" id="table1">
  <tr>
    <td background="images/top_middle.jpg">
    <img src="images/top_left.jpg" width="423" height="91" alt="" ></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif"><img src="images/top_right.jpg" width="365" height="91" alt=""></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" id="table2">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif"></td>
  </tr>
</table>
<p align="right"><font class="copyright">Copyright (c) 2001-<%=Year(Now)%>, Rate-Highway, Inc. All Rights Reserved.</font></p>
<p>
<p align="center"><font size="5" color="#384F5B">We are sorry to report that an 
error has occurred.<br>
&nbsp;</font></p><p>Please print and fax this page to 
Rate-Highway Support at (888) 551-0029</p>
<p><u><font size="2">Debug information&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
</u></p>
<p>
<%=strErrorListing %>
<br>
<font face="Courier New">
			   <%	For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"

       
					Next
				%>
</font></p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
