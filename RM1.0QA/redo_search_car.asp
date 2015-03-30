<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 

<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 0
   
   'On Error Resume Next
 
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
	Dim Whatever
	Dim strDowList

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	If IsNumeric(strUserId) = False Then
		strUserId = 0
	End If
	
	intArray = Split(Request.Form("shop_request_id"), ",")
	intCount = UBound(intArray)

	strConn = Session("pro_con")
	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_shop_request_rerequest"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id"   , adInteger, 1, 0 )
	
	While intCount >= 0
		adoCmd.Parameters("@shop_request_id").Value = intArray(intCount)

		adoCmd.Execute 
		intCount = intCount - 1
	
	Wend
	
	Set adoCmd = Nothing	
	
	Server.Transfer "search_queue_car.asp"	
	
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>redo car search</title>
</head>
<body>

			   <%	For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "</b> <br>"

       
					Next
				%>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
