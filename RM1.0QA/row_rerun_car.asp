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
	Rem we have no clue how many, so cross your fingers
	Dim Whatever

	
'	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	
	strConn = Session("pro_con")
	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_shop_request_search_id"
	adoCmd.CommandType = 4  

	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, Request("shop_request_id"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 5, Request("city_cd"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cd", 200, 1, 4, Request("shop_car_type_cd"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@data_source", 200, 1, 3, Request("data_source"))
    adoCmd.Parameters.Append adoCmd.CreateParameter("@begin_arv_dt", 135, 1, 0, Request("arv_dt")) 'FormatDateTime(Request("arv_dt")))
 
	If Request("debug") = "true" Then
	
	Else
		Set adoRS = adoCmd.Execute
		

		If adoRS.EOF = False Then
			Set adoCmd = CreateObject("ADODB.Command")

			adoCmd.ActiveConnection =  strConn
			adoCmd.CommandText = "car_shop_request_search_update"
			adoCmd.CommandType = 4
		
			adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_search_id", 3, 1, 0, adoRS.Fields("shop_request_search_id").Value)
			adoCmd.Parameters.Append adoCmd.CreateParameter("@request_status", 200, 1, 1, "N")
			adoCmd.Parameters.Append adoCmd.CreateParameter("@request_msg", 200, 1, 99, "")

			adoCmd.Execute
			
		End If
		
	End If

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Row re-run request</title>
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
</head>
<body >
	<p align="center"><font face="Tahoma">Report Row re-requested successfully.</font></p>
    <p align="center"><font face="Tahoma"><a href="javascript:window.close()">Close</a></font></p>
    <p align="center">&nbsp;</p>
    <p align="center">&nbsp;</p>
<p align="center"><font face="Tahoma" size="1">id: <%=adoRS.Fields("shop_request_search_id").Value%></font></p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>