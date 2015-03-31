<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

	On Error Resume Next

   Server.ScriptTimeout = 180
 
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
	Dim varCarTypes()
	Dim varDataSources()
	Dim varCities()
	Dim varVendors()
	Dim varDates()
	Dim strSelectedVendor
	Dim strBgColor
	Dim blnDarkRow 
	Dim curRate
	Dim curTotal
	Dim strCarList
	Dim strVendList
	Dim strCarCodeListArray 
	Dim strCarCodeList
	Dim strDowString
	Dim blnRedoEnabled
	Dim strIPAddress
	Dim strLOR	
	Dim IsMultiRate
	Dim strRateColumn
	Dim intDisplayRateType 
	Dim blnOnewayReverse 
	
		
	strIPAddress = Request.Servervariables("REMOTE_ADDR") 

	If UCASE(Request("redoenabled")) = "TRUE" Then
		blnRedoEnabled = True
	Else
		blnRedoEnabled = False
	End If

	intReportRequestId = Request("reportrequestid")
	strSecurityCode =    Request("security_code")
	strCarTypeCd =       Request("car_type_cd")
	strCityCd =          Request("city_cd")
	intDisplayRateType = Request("displayratetype")
	
	If (intReportRequestId < 10000000) Then
		strConn = Session("pro_con")
	ElseIf (intReportRequestId > 10000000) And (intReportRequestId < 19999999) Then
		strConn = Session("pro_con_vanguard")
	ElseIf (intReportRequestId > 20000000) And (intReportRequestId < 29999999) Then
		strConn = Session("pro_con")
	ElseIf (intReportRequestId > 30000000) And (intReportRequestId < 39999999) Then
		strConn = Session("pro_con_thor")
	End If
	
  	Set adoRS = CreateObject("ADODB.Recordset")
  	Set adoCmd = CreateObject("ADODB.Command")
	adoRS.CursorLocation = adUseClient
	
	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "car_shopped_rate_select_rpt2"
	adoCmd.CommandType = 4
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, intReportRequestId) 
	adoCmd.Parameters.Append adoCmd.CreateParameter("@report", 3, 1, 0, 1)
	
	If Len(Request.QueryString("vend_override")) = 2 Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_override", 200, 1, 2, Request.QueryString("vend_override"))
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_override", 200, 1, 2, Null)
	End If

	adoCmd.Parameters.Append adoCmd.CreateParameter("@ipaddress", 200, 1, 20, strIPAddress)

	If strCityCd = "ALLLL" And strCarTypeCd = "ALLL" Then
		strCarTypeCd = ""
	End If	
		
	If strCarTypeCd = "" Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cd", 200, 1, 4, Null)	
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cd", 200, 1, 4, strCarTypeCd)
	End If
	
	If strCityCd = "" Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 6, Null)	
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 6, strCityCd)
	End If

	adoRS.Open adoCmd, , adOpenStatic, adLockReadOnly, adCmdStoredProc 

	Dim intDateIndex
	Dim intCarTypeIndex
	Dim intDataSourceIndex
	Dim intCityIndex
	
	If adoRS.State = adStateClosed Then
	   Rem If the report request was bogus
	   Server.Transfer "invalid_report.asp"
	
	End If

	
	
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


	%>
<html>

<head>
<link rel="SHORTCUT ICON" href="http://www.rate-highway.com/favicon.ico">
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor.com | View Report By Car Type</title>
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<style>
<!--
.report_detail_dark { 	height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left; font-weight:bold  }
.report_detail_light { 	height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left}
.copyright {	FONT-SIZE: 0.7em; TEXT-ALIGN: right}
body         { font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt }

-->
</style>
</head>
<body>
<p><font style="font-size: 7pt">debug: <%=intDisplayRateType %> / <%=IsNumeric(Request("display_rate_type")) %> </font> </p>
<%
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

    Set adoRS = Nothing 
    Set adoCmd = Nothing 
   
%>
</body>

</html>