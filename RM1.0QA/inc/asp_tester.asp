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
	Dim varHotels(5)
	Dim varHotelIds(5)
	Dim varSites(9)
	Dim varSiteIds(9)
	Dim intResults
	Dim intPrice
	Dim varDates(30)
	Dim intIndex
	Dim blnExit
	Dim intReportRequestID
	Dim intErrorCount
	Dim intTimeoutCount

	
	'intReportRequestID = Request("ReportRequestID")
	
	'strConn = Session("pro_con")
	
  	'Set adoCmd = CreateObject("ADODB.Command")

	'adoCmd.ActiveConnection = strConn
	'adoCmd.CommandText = "symSearchResultSelect30dayByHotel"
	'adoCmd.CommandType = 4

	'adoCmd.Parameters.Refresh 

	'adoCmd.Parameters("@ReportRequestID").Value = intReportRequestID 

	intReportRequestID = 47715 'Request("ReportRequestID")
	
	strConn = Session("pro_con")
	

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>asp tester</title>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
</head>

<body>

<p>Connection = <%=strConn %></p>

</body>

</html>