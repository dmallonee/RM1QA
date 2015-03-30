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
	Dim debug
	Dim strArray
	Dim intArrayHolder
	Dim strRateArray
	Dim strRateString
	Dim strErrorListing
	
	strErrorListing = ""

    strUserId =    Request.Cookies("rate-monitor.com")("user_id")
	strCityCd =    Request.Form("city_cd")
	strCarTypeCd = Request.Form("car_type_cd")
	
	
	If strCityCd = "" Or strCarTypeCd = "" Then
		strErrorListing = "Your session has expired, please logout and back in"

	Else 

		strConn = Session("pro_con")
	
		debug = Request.Form("debug")
		
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "utilization_car_group_delete"
		adoCmd.CommandType = adCmdStoredProc

		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",      3, 1,  0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",     200, 1, 6, strCityCd)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd", 200, 1, 4, strCarTypeCd)

		Call adoCmd.Execute(,,adExecuteNoRecords)
	
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

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "utilization_car_group_insert"
		adoCmd.CommandType = adCmdStoredProc

		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",           3, 1, 0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd",         200, 1, 6, strCityCd)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@grp_car_type_cd", 200, 1, 4, strCarTypeCd)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd",     200, 1, 4)
		
		strArray = Split(Request.Form("car_type_list"), ",", -1, 1)
	
		For intArrayHolder = LBound(strArray) To UBound(strArray)
	
			If (debug = 1) Then
				Response.Write "DEBUG - " & intArrayHolder & " - " & strArray(intArrayHolder) & ", " &  strRateArray(intArrayHolder) & "<BR>"
			End If

			adoCmd.Parameters("@car_type_cd").Value = strArray(intArrayHolder)

			Call adoCmd.Execute(,,adExecuteNoRecords)

		Next

		Set adoCmd = Nothing
		
	End If
	
	If (err.number <> 0) Or (debug = 1) Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   strErrorListing = strErrorListing & pad & "<b>Error adding rate change to queue</b><br>"
	   strErrorListing = strErrorListing & pad & "Error Number   = #<b>" & err.number & "</b><br>"
	   strErrorListing = strErrorListing & pad & "Error Desc.    = <b>" & err.description & "</b><br>"
	   strErrorListing = strErrorListing & pad & "Help Context   = <b>" & err.HelpContext & "</b><br>"
	   strErrorListing = strErrorListing & pad & "Help File Path = <b>" & err.helpfile & "</b><br>"
	   strErrorListing = strErrorListing & pad & "Error Source   = <b>" & err.source & "</b><br><hr>"

	ElseIf strErrorListing <> "" Then
		Rem don't change pages
		
	Else
		Server.Transfer "system_utilization_car_groups.asp"
	
	End If


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>search_request_insert</title>
<style>
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
    <img src="images/top_left.jpg" width="423" height="91"></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif"><img src="images/top_right.jpg" width="365" height="91"></td>
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
