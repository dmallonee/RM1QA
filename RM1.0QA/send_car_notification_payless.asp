<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 

<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 0
 
	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoPrices
	Dim intIndex
	Dim blnExit
	Dim intReportRequestID
	Dim intErrorCount
	Dim intTimeoutCount
	Dim strCarType 
	Dim intResults
	Dim intPrice

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_notify_user_payless"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request.Form("search_id"))
	
	adoCmd.Execute
	
	'Server.Transfer "search_queue_car_payless.asp"	
	
	Response.Redirect "search_queue_car_payless.asp"	

	Set adoCmd = Nothing


%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>search_request_insert</title>
</head>
<body>

			   <%	For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"

       
					Next
				%>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
