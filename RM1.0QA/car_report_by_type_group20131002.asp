<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<!-- #INCLUDE FILE="inc/DanDate.inc"-->
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
	Dim varCurrency()	
	Dim strSelectedVendor
	Dim strBgColor
	Dim blnDarkRow 
	Dim curRate
	Dim curTotal
	Dim strCarList
	Dim strVendList
	Dim strCurrencyList	
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
	Dim intLocale
	
	intUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	If IsNumeric(intUserId) = False Then
		intUserId = 1
		
	ElseIf intUserId < 1 Then
		intUserId = 1
		
	End If
			
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
	strVendor          = Request("vend_override")
	
	If (intReportRequestId < 10000000) Then
		strConn = Session("pro_con")
	ElseIf (intReportRequestId > 10000000) And (intReportRequestId < 19999999) Then
		strConn = Session("pro_con_vanguard")
	ElseIf (intReportRequestId > 20000000) And (intReportRequestId < 29999999) Then
		strConn = Session("pro_con")
	ElseIf (intReportRequestId > 30000000) And (intReportRequestId < 39999999) Then
		strConn = Session("pro_con_thor")
	End If
	
  	Set adoCmd = CreateObject("ADODB.Command")
  	
	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "currency_select"
	adoCmd.CommandType = 4
	
	Set adoRCurrency = adoCmd.Execute

  	Set adoCmd = CreateObject("ADODB.Command")
  	
	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "user_rate_rpt_select"
	adoCmd.CommandType = 4
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0, intUserId) 

	Set adoRSettings = adoCmd.Execute
	
	If adoRSettings.Fields("us_decimal").Value Then
		intLocale = SetLocale(1033) ' US
	Else
		intLocale = SetLocale(1031) ' Germany
	End If
  	
  	
  	Set adoRS = CreateObject("ADODB.Recordset")
  	Set adoCmd = CreateObject("ADODB.Command")
	adoRS.CursorLocation = adUseClient
	
	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "car_shopped_rate_select_rpt_group"
	adoCmd.CommandType = 4
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, intReportRequestId) 
	adoCmd.Parameters.Append adoCmd.CreateParameter("@report", 3, 1, 0, 1)
	
	If Len(strVendor) = 2 Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_override", 200, 1, 2, strVendor)
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

	While NOT adoRS.EOF 
		intDateIndex = adoRS.Fields("date_count").Value
		intDataSourceIndex = adoRS.Fields("data_source_count").Value
		intCarTypeIndex =    adoRS.Fields("car_type_count").Value
		intVendorIndex =     adoRS.Fields("vendor_count").Value
		strCarCodeList =     adoRS.Fields("car_type_list").Value
		strDowString =       adoRS.Fields("dow_list").Value	
		strLOR =             adoRS.Fields("lor").Value	
		intCityIndex =       adoRS.Fields("city_count").Value	
		strCityCodeList =    adoRS.Fields("city_cd_list").Value
		If IsNumeric(intDisplayRateType) = False Then
			Rem Allow the user to override the database setting
			intDisplayRateType = adoRS.Fields("display_rate_type").Value
		End If
		strAlertStatus =     adoRS.Fields("alert_status").Value
		intRateChanges =     adoRS.Fields("rate_changes").Value	
		tmpDisplayRateType = adoRS.Fields("display_rate_type").Value
		blnOnewayReverse   = adoRS.Fields("oneway_reverse").Value
		strReportName      = adoRS.Fields("rpt_desc").Value
	
		adoRS.MoveNext

	Wend


	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Errors Occured! - Stage 2 of Report Load<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If


	strCarCodeListArray = Split(strCarCodeList, ",")
	strCityCodeListArray = Split(strCityCodeList, ",")

	ReDim varCarTypes(intCarTypeIndex)
	ReDim varDataSources(intDataSourceIndex)
	ReDim varDates(intDateIndex)
	ReDim varVendors(intVendorIndex)
	ReDim varVendorCds(intVendorIndex)
	ReDim varCities(intCityIndex)
	
	intDateIndex = 0
	intCarTypeIndex = 0
	intDataSourceIndex = 0
	intVendorIndex = 0
	intCityIndex = 0
	
	Set adoRS = adoRS.NextRecordset
	
	While adoRS.EOF = False
	
	   If intDateIndex < UBound(varDates) Then
		   If varDates(intDateIndex) <> adoRS.Fields("arv_dt").Value Then
			   intDateIndex= intDateIndex+ 1
	    	   If intDateIndex <= UBound(varDates) Then
					varDates(intDateIndex) = adoRS.Fields("arv_dt").Value
				End If
			End if
		End If
		
		If intCarTypeIndex < UBound(varCarTypes) Then
			If varCarTypes(intCarTypeIndex ) <> adoRS.Fields("group_car_type").Value Then
				intCarTypeIndex = intCarTypeIndex + 1
				If intCarTypeIndex <= UBound(varCarTypes) Then
					varCarTypes(intCarTypeIndex ) = adoRS.Fields("group_car_type").Value 
					strCarList = strCarList & adoRS.Fields("group_car_type").Value & ", " 
				End If
			End If
		End If

		If intDataSourceIndex < UBound(varDataSources) Then
			If varDataSources(intDataSourceIndex) <> adoRS.Fields("data_source_name").Value Then
				intDataSourceIndex = intDataSourceIndex + 1
				If intDataSourceIndex <= UBound(varDataSources) Then
					varDataSources(intDataSourceIndex) = adoRS.Fields("data_source_name").Value
					strSourceList = strSourceList & adoRS.Fields("data_source_name").Value & ", " 
				End If
			End If
		End If

		If intVendorIndex < UBound(varVendors) Then
			If varVendors(intVendorIndex) <> adoRS.Fields("vendor_name_rpt").Value Then
				intVendorIndex = intVendorIndex + 1
				If intVendorIndex <= UBound(varVendors) Then
					varVendors(intVendorIndex) = adoRS.Fields("vendor_name_rpt").Value
					varVendorCds(intVendorIndex) = adoRS.Fields("vend_cd").Value
					strVendList = strVendList & adoRS.Fields("vendor_name_rpt").Value & ", " 
				End If
			End If
		End If

		If intCityIndex < UBound(varCities) Then
			If varCities(intCityIndex ) <> adoRS.Fields("city_cd").Value Then
				intCityIndex = intCityIndex + 1
				If intCityIndex <= UBound(varCities) Then
					If (adoRS.Fields("city_cd").Value <> adoRS.Fields("rtrn_city_cd").Value) And (IsNull(adoRS.Fields("rtrn_city_cd").Value) = False) Then
						varCities(intCityIndex ) = adoRS.Fields("city_cd").Value & "-" & adoRS.Fields("rtrn_city_cd").Value
						strCityList = strCityList & adoRS.Fields("city_cd").Value & "-" & adoRS.Fields("rtrn_city_cd").Value & ", " 
					Else
						varCities(intCityIndex ) = adoRS.Fields("city_cd").Value
						strCityList = strCityList & adoRS.Fields("city_cd").Value & ", " 
					End If
				End If
			End If
		End If

		adoRS.MoveNext
			
	Wend


	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Errors Occured! - Stage 3 of Report Load<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If


	Rem Allow the user to override this
	If IsNumeric(Request("display_rate_type")) Then
		If Request("display_rate_type") > 0 And Request("display_rate_type") < 5 Then
			intDisplayRateType = Request("display_rate_type")
		Else
			Rem otherwise leave it as is from teh database
		End If
	
	End If

	Dim strRateTitle
	Dim strRateString

	Rem
	Rem The user may have requested this report form the search queue, in which case the 
	Rem intDisplayRateType will be blank. If so, grab the value from the temp holder 
	Rem

	If intDisplayRateType = "" Then
		intDisplayRateType = tmpDisplayRateType 
	End If

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Errors Occured! - Stage 4 of Report Load<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If


	Select Case intDisplayRateType 
	
		Case 2
			IsMultiRate = False
			strRateColumn = "total_rt_amt"
			strRateTitle = "(total rate)"
			strRateString = "Total rate amount (less tax and fees)"
			
		Case 3
			IsMultiRate = False
			strRateColumn = "est_rental_chrg_amt"
			strRateTitle = "(total price)"
			strRateString = "Total Price (includes tax and fees)"

		Case 4
			IsMultiRate = True
			strRateColumn = "est_rental_chrg_amt"
			strRateTitle = "(base rate/total price)"
			strRateString = "Rate amount & Total Price"

		Case 5
			IsMultiRate = True
			strRateColumn = "est_rental_chrg_amt"
			strRateTitle = "(base rate/total price/drop chg)"
			strRateString = "Rate amount & Total Price & Drop Charge"

		Case 6
			IsMultiRate = True
			strRateColumn = "est_rental_chrg_amt"
			strRateTitle = "(base rate/total price/extra day)"
			strRateString = "Rate amount & Total Price & Extra day"

		Case 8
			IsMultiRate = False
			strRateColumn = "est_rental_chrg_amt"
			strRateTitle = "(total price)"
			strRateString = "Total Price (includes tax and fees)"

		Case 9
			IsMultiRate = False
			strRateColumn = "est_rental_chrg_amt"
			strRateTitle = "(total price)"
			strRateString = "Total Price (includes tax and fees)"


		Case Else
			IsMultiRate = False
			strRateColumn = "rt_amt"
			strRateTitle = "(base rate)"
			strRateString = "Rate amount (daily or weekly rate)"

	End Select

	adoRS.MoveFirst
	
	Rem Added a third result set in the sp because I can't get the MoveFirst to not error out
	'Set adoRS = adoRS.NextRecordset
	
	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Errors Occured! - Final Stage of Report Load<br>"
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
<title>Rate-Monitor.com | Rate Report by Group</title>
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<script language='Javascript' type="text/javascript" > 

	var userLang = (navigator.language) ? navigator.language : navigator.userLanguage; 
	//alert ("The language is: " + userLang);

	function centerPopUp( url, name, width, height, scrollbars ) { 
 
	if( scrollbars == null ) scrollbars = "0" 
 
	str  = ""; 
	str += "resizable=1,"; 
	str += "scrollbars=" + scrollbars + ","; 
	str += "width=" + width + ","; 
	str += "height=" + height + ","; 
    
	if ( window.screen ) { 
		var ah = screen.availHeight - 30; 
		var aw = screen.availWidth - 10; 
 
		var xc = ( aw - width ) / 2; 
		var yc = ( ah - height ) / 2; 
 
		str += ",left=" + xc + ",screenX=" + xc; 
		str += ",top=" + yc + ",screenY=" + yc; 
	} 
	window.open( url, name, str ); 
} 

</script> 
<style type="text/css" >
<!--
.report_detail_dark { 	height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left; font-weight:bold  }
.report_detail_light { 	height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left}
.copyright {	FONT-SIZE: 0.7em; TEXT-ALIGN: right}
body         { font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt }

-->
</style>
</head>

<body topmargin="0">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_middle.jpg">
    <img src="images/top_left.JPG" width="424" height="91"></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif"><img src="images/top_right.jpg" width="365" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif"></td>
  </tr>
</table>
<p align="right"><font class="copyright">Copyright � 2001-<%=Year(Now)%>,
<a target="_blank" href="http://www.rate-highway.com">Rate-Highway, Inc.</a> (www.rate-highway.com) 
All Rights Reserved.<br>
Rate-Monitor is a product of Rate-Highway, Inc. - the creators of the first 
fully automated rate positioning tool.</font></p>


	
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#92393A" width="100%" id="AutoNumber2">
      <tr>
        <td width="100%" bgcolor="#92393A">&nbsp;<font color="#FFFFFF" size="4" face="Verdana">Rate Detail 
		- <%=strReportName %></font></td>
      </tr>
</table>


<table cellSpacing="0" cellPadding="8" width="100%" border="0">
  <tr>
    <td bgColor="#CFD7DB" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif" width="605">
		<form method="GET" name="rate_detail">

			
                  <select size="1" name="city_cd" title="City to display">
				  <% For intIndex = LBound(strCityCodeListArray) To UBound(strCityCodeListArray)	%>
				  	<% If strCityCd = strCityCodeListArray(intIndex) Then             %>
				   				   <option selected><%=strCityCodeListArray(intIndex) %> </option>
				  	<% Else             %>
				   				   <option ><%=strCityCodeListArray(intIndex) %> </option>
				  	<% End If           %>

				  <% Next %>  				 
 
			      <% If strCityCd = "ALLLL" Then %>
				   <option selected value="ALLLL" >ALL </option>
                  <% Else %>
				   <option value="ALLLL" >ALL </option>
                  <% End If %>	

 
		           </select>,
                  <select size="1" name="car_type_cd" title="Car type to display">
				  <% For intIndex = LBound(strCarCodeListArray) To UBound(strCarCodeListArray)	%>
				  <% If strCarTypeCd = strCarCodeListArray(intIndex) Then             %>
				   				   <option value="<%=strCarCodeListArray(intIndex) %>" selected><%=strCarCodeListArray(intIndex) %> </option>
				  <% Else             %>
				   				   <option value="<%=strCarCodeListArray(intIndex) %>" ><%=strCarCodeListArray(intIndex) %> </option>
				  <% End If           %>
 
				  <% Next %>                  
                  
                  
                  <% If strCarTypeCd = "ALLL" Then %>
				   <option selected value="ALLL" >ALL </option>
                  <% Else %>
				   <option value="ALLL" >ALL </option>
                  <% End If %>	
                  </select>, 
                  <% If adoRSettings.Fields("show_vendor").Value Then %>
                  <select name="vend_override" title="Comparison vendor">
                  <% For intVendorIndex = LBound(varVendors)  To UBound(varVendors) %>
                    <% If Len(varVendorCds(intVendorIndex)) = 2 Then %>
						<% If varVendorCds(intVendorIndex) = strVendor Then %>
					  		<option selected="selected" value="<%=varVendorCds(intVendorIndex)%>"><%=varVendors(intVendorIndex)%></option>
					  	<% Else %>
					  		<option value="<%=varVendorCds(intVendorIndex)%>"><%=varVendors(intVendorIndex)%></option>
					  	<% End If %>
				  	<% End If %>
				  <% Next %>
				  </select>,
				  <% End If %> 
                  <% If adoRSettings.Fields("show_currency").Value Then %>
                  <select name="currency_override" title="Display currency">
					   <optgroup label="North American">
					     <option value="USD">USD</option>
					     <option value="CAD">CAD</option>
					   </optgroup>
					   <optgroup label="European">
					     <option value="EUR">EUR</option>
					     <option value="GBP">GBP</option>
					   </optgroup>
				  </select>,
				  <% End If %> 
				  <select size="1" name="displayratetype" style="width:225" title="Rate type to display">
                    <% Select Case intDisplayRateType %>
                    <%    Case 2   %>
                    <option value="1">Base rate amt</option>
                    <option value="2" selected>Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <option value="7">Rate amt/Total price/Limit/Extra</option>
                    <option value="8">Total price/Ins. Included</option>
                    <option value="9">Total price/Ins. Not Included</option>

                    <%    Case 3   %>
                    <option value="1">Base rate amt</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3" selected >Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <option value="7">Rate amt/Total price/Limit/Extra</option>
                    <option value="8">Total price/Ins. Included</option>
                    <option value="9">Total price/Ins. Not Included</option>
                    
                    <%    Case 4   %>
                    <option value="1">Base rate amt</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4" selected >Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <option value="7">Rate amt/Total price/Limit/Extra</option>
                    <option value="8">Total price/Ins. Included</option>
                    <option value="9">Total price/Ins. Not Included</option>

                    <%    Case 5   %>
                    <option value="1">Base rate amt</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5" selected >Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <option value="7">Rate amt/Total price/Limit/Extra</option>
                    <option value="8">Total price/Ins. Included</option>
                    <option value="9">Total price/Ins. Not Included</option>

                    <%    Case 6   %>
                    <option value="1">Base rate amt</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6" selected>Rate amt/Total price/Extra day</option>
                    <option value="7">Rate amt/Total price/Limit/Extra</option>
                    <option value="8">Total price/Ins. Included</option>
                    <option value="9">Total price/Ins. Not Included</option>

                    <%    Case 7   %>
                    <option value="1">Base rate amt</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <option value="7" selected>Rate amt/Total price/Limit/Extra</option>
                    <option value="8">Total price/Ins. Included</option>
                    <option value="9">Total price/Ins. Not Included</option>

                    <%    Case 8  %>
                    <option value="1">Base rate amt</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <option value="7">Rate amt/Total price/Limit/Extra</option>
                    <option value="8" selected >Total price/Ins. Included</option>
                    <option value="9">Total price/Ins. Not Included</option>

                    <%    Case 9  %>
                    <option value="1">Base rate amt</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <option value="7">Rate amt/Total price/Limit/Extra</option>
                    <option value="8">Total price/Ins. Included</option>
                    <option value="9" selected >Total price/Ins. Not Included</option>

                    <%    Case Else   %>
                    <option value="1" selected >Base rate amt</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <option value="7">Rate amt/Total price/Limit/Extra</option>
                    <option value="8">Total price/Ins. Included</option>
                    <option value="9">Total price/Ins. Not Included</option>

                    <% End Select %>
                    </select>
                  <input type="submit" value="Display" name="display" class="rh_button" id="display" style="border:3px double #2A2F34; font-family: Vendana, Arial, Helvetica, sans-serif; font-size: 10pt; color:#FFFFFF; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1; background-color:#384F5B; font-weight:bold" ><input type="hidden" name="reportrequestid" value="<%=Request("reportrequestid") %>"><input type="hidden" name="redoenabled" value="<%=Request("redoenabled") %>"> </form>    
    </td>
    <td vAlign="top" bgColor="#CFD7DB" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif" >
    </td>
  </tr>
</table>

	
	<%



	While NOT adoRS.EOF 

		If (strCarType <> adoRS.Fields("group_car_type").Value) Or _
		   ((strCityCd <> adoRS.Fields("city_cd").Value & " to " & adoRS.Fields("rtrn_city_cd").Value & "") And (strCityCd <> adoRS.Fields("city_cd").Value)) Then
			
			strCarType = adoRS.Fields("group_car_type").Value
			If (IsNull(adoRS.Fields("rtrn_city_cd").Value)) Or (adoRS.Fields("city_cd").Value = adoRS.Fields("rtrn_city_cd").Value & "") Then
				strCityCd = adoRS.Fields("city_cd").Value 
			' Not necessary to reverse - already done in the db	
			'ElseIf CBool(blnOnewayReverse) Then
			'	strCityCd = adoRS.Fields("rtrn_city_cd").Value & " to " & adoRS.Fields("city_cd").Value
			Else
				strCityCd = adoRS.Fields("city_cd").Value & " to " & adoRS.Fields("rtrn_city_cd").Value
			End If

			%>
			<table cellSpacing='0' cellPadding='8' width='100%' border='0'>
			  <tr>
			    <td vAlign="top" bgColor="#CFD7DB" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">    	
                  <p><font size="+1"><b><%=strCityCd %></b></font>, <font size="+1"><b><%=strCarType %></b><%=strRateTitle %>       
					</font><table cellSpacing="0" cellPadding="2" border="1">
		      <tr>
			<!--
		        <th align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
		        &nbsp;</th>
			-->
		      	<%
		      	For intVendorIndex = LBound(varVendors)  To UBound(varVendors) 
		      		blnMoved = False
		      	%>
		        <th align="middle" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=varVendors(intVendorIndex) %>&nbsp;</th>
		      	<%
		      	Next
		      	%>

				<% If blnRedoEnabled Then %> 
		        <th align="middle" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
                				Redo</th>
				<% End If %>
<!--		        <th noWrap align="middle" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
		        Details</th>
-->		        
		      </tr>
		      <%
		      
		      %>
		      <tr>
		      <% 
		      	blnDarkRow = True
		      	
				For intIndex = 1 To intDateIndex 'LBound(varDates) To UBound(varDates)	- 1	      
					blnMoved = False
				
					'Dim strSelectedVendor
					
					blnDarkRow = Not blnDarkRow
					
					If adoRS.EOF = True Then
						exit for
					end if
					
					curRate = adoRS.Fields(strRateColumn).Value
					curTotal = adoRS.Fields("est_rental_chrg_amt").Value				

		      %>
		      <% If adoRSettings.Fields("us_date").Value Then %>
		      <th noWrap align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=WeekDayName(Weekday(adoRS.Fields("arv_dt").Value ), True) & " - " & FormatDateTime(adoRS.Fields("arv_dt").Value , 2) %>&nbsp;</th>
		      <% Else %>
		      <th noWrap align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=DanDate(adoRS.Fields("arv_dt").Value , "%d/%m/%y") %>&nbsp;</th>
		      <% End If %>
		      
			  
		      <% 
				For intVendorIndex = LBound(varVendors) To UBound(varVendors) '- 1	
				
					'If adoRS.Fields("data_source_name").Value = "Hertz" Then
					'	strBgColor = "#FFFFFF"
					'Else
						If blnDarkRow Then
							strBgColor = "#B2BEC4"
						Else
							strBgColor = "#CFD7DB"
						End If					
					'End If	   
					
					If adoRS.EOF = True Then
						Exit For
					End If   
					
					
					If (intVendorIndex = UBound(varVendors)) Then
		
						If blnRedoEnabled Then
					      %>
		    			    <td noWrap align="center" bgColor="#B2BEC4" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif">
                  <!--          <a href="javascript:centerPopUp( 'row_rerun_car.asp?shop_request_id=<%=adoRS.Fields("shop_request_id").Value%>&city_cd=<%=adoRS.Fields("city_cd").Value%>&group_car_type=<%=adoRS.Fields("group_car_type").Value%>&data_source=<%=adoRS.Fields("data_source").Value%>&arv_dt=<%=Server.URLEncode(adoRS.Fields("arv_dt").Value)%>', 're-request', 400, 250, 1 )">
                  -->
                            <a href="javascript:centerPopUp( 'row_rerun_car.asp?shop_request_id=<%=adoRS.Fields("shop_request_id").Value%>&city_cd=<%=adoRS.Fields("city_cd").Value%>&shop_car_type_cd=<%=adoRS.Fields("group_car_type").Value%>&data_source=<%=adoRS.Fields("data_source").Value%>&arv_dt=<%=Server.URLEncode(adoRS.Fields("arv_dt").Value)%>', 'rerequest', 400, 250, 0 );" >
                            <img alt="" src="images/re_run.jpg" align="middle" width="18" height="18" border="0"></a></td>
						  <%
						End If
		      
		      		ElseIf IsNumeric(adoRS.Fields(strRateColumn).Value) Then
		      		
			      		If adoRS.Fields("rent_mi_alwnc_cd").Value = "Y" Then
			      			strMilage = "<sup>" & adoRS.Fields("mi_km_ind").Value & "</sup>"
			      		Else
			      			strMilage = "&nbsp;"
			      		End If
			      		
			      		If Len(adoRS.Fields("extra_values").Value) > 0  Then
			      		    strInsurance = Replace(adoRS.Fields("extra_values").Value, "|", ",")
			      		    strInsurance = Left(strInsurance, (Len(strInsurance) - 1)) 
			      			strInsurance = "<sup>" & strInsurance & "</sup>"
			      		Else
			      			strInsurance = "&nbsp;"
			      		End If			 
			 

		      			If adoRS.Fields(strRateColumn).Value = -1 Then
		      				Rem As in not searched - no rate should be displayed
					      %>
		    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;</td>
						  <%


		      			ElseIf curRate > adoRS.Fields(strRateColumn).Value Then
		      			
		      				rem If curTotal = adoRS.Fields("est_rental_chrg_amt").Value
					      %>
					      	<% Select Case intDisplayRateType %>
					      		<% Case 4 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 5 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & FormatNumber("0" & adoRS.Fields("extra_values").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 6 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & FormatNumber("0" & adoRS.Fields("extra_da_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 7 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & adoRS.Fields("rent_mi_alwnc").Value & "/" & adoRS.Fields("mi_chrg_amt").Value %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 8 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %>
				    			    <%=strInsurance %></td>

				      			<% Case Else %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %>
				    			    <sub><%=trim(adoRS.Fields("shop_car_type_cd").Value) %></sub></td>
			      			<% End Select %>
						  <%
						  
						ElseIf curRate < adoRS.Fields(strRateColumn).Value Then
						
					      %>
					      	<% Select Case intDisplayRateType %>
					      		<% Case 4 %>
							        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 5 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & FormatNumber("0" & adoRS.Fields("extra_values").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 6 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & FormatNumber("0" & adoRS.Fields("extra_da_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 7 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & adoRS.Fields("rent_mi_alwnc").Value & "/" & adoRS.Fields("mi_chrg_amt").Value %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 8 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %>
				    			    <%=strInsurance %></td>

					      		<% Case Else %>
							        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %>
    				    			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value) %></sub></td>
				      		<% End Select %>
						  <%
						  
						Else
					      %>
					      	<% Select Case intDisplayRateType %>
					      		<% Case 4 %>
							        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 5 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & FormatNumber("0" & adoRS.Fields("extra_values").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 6 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & FormatNumber("0" & adoRS.Fields("extra_da_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 7 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & adoRS.Fields("rent_mi_alwnc").Value & "/" & adoRS.Fields("mi_chrg_amt").Value %>
					      			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value & " " & adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 8 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %>
				    			    <%=strInsurance %></td>

					      		<% Case Else %>
							        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %>
    				    			<sub><%=trim(adoRS.Fields("shop_car_type_cd").Value) %></sub></td>

				      		<% End Select %>
						  <%
						
						
						End If						
						

					Else
					
					
	      				If adoRS.Fields("shop_msg").Value = "Rate was unavailable when website attempted to confirm with GDS" Then
		      				Rem As in not searched - no rate should be displayed
					      %>
		    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" title="<%=adoRS.Fields("shop_msg").Value %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif">GDS Link</td>
						  <%

						Else

					      %>
					        <td noWrap align="right" bgColor="<%=strBgColor  %>" title="<%=adoRS.Fields("shop_msg").Value %>" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">
			                <sup> </sup>Closed</td>
						  <%

						End If

					End If


					Rem we need the details from the current date, prior to moving on to the next
					Rem date batch, so donn't advance after the last rate, but keep the ado row around
					Rem until the last coloumn can be created and then advance.
					If  intVendorIndex <> UBound(varVendors) - 1 Then
					
				  		If adoRS.EOF = False Then
						    adoRS.MoveNext
						    blnMoved = True
						End If
						
					End If
					
	

			    Next

		  		'If adoRS.EOF = False AND blnMoved = False Then
				'    adoRS.MoveNext
				'    blnMoved = True
				'End If

				blnMoved = False

			  %>
<!--			        
		        <td noWrap bgColor="<%=strBgColor%>" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">Lowest</td>
-->
		      </tr>
		      <%
		      	Next 
		      %>
		      
      		</table>
      		</p></td>
      	</table>
			<%
			
		End If
		
		'adoRS.MoveNext
		
	Wend
	
	adoRS.MoveFirst
'
'	'While (adoRS.EOF = False) And (adoRS.Fields(strRateColumn).Value = -1)
'		adoRS.MoveNext
'
'	Wend


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

'SUB ErrorADOReport(parm_msg,parm_conn)
'   HowManyErrs=parm_conn.errors.count
'   IF HowManyErrs=0 then
'      exit sub
'   END IF
'   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
'   response.write "<b>ADO Reports these Database Error(s) executing:<br>"
'          response.write SQLstmt & "</b><br>"
'   for counter= 0 to HowManyErrs-1
'      errornum=parm_conn.errors(counter).number
'      errordesc=parm_conn.errors(counter).description
'      response.write pad & "Error#=<b>" & errornum & "</b><br>"
'      response.write pad & "Error description=<b>"
'      response.write errordesc & "</b><p>"
'   next
'END SUB

	 	
 %>


<table cellSpacing="0" cellPadding="8" width="100%" border="0">
  <tr>
    <td vAlign="top" bgColor="#CFD7DB" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">

    <p>
    <% If strAlertStatus = "S" Then %>
	    <a href="rate_change_report.asp?reportrequestid=<%=intReportRequestId %>&amp;security_code=<%=Escape(strSecurityCode) %>" name="rate_change" target="_blank">Change my rates</a> using this report
       <br>
	    View my <a  target="_blank" href="rate_change_receipt.asp?reportrequestid=<%=intReportRequestId %>&security_code=<%=Escape(strSecurityCode) %>">rate change requests</a> from this report
    <% ElseIf strAlertStatus = "P" Then %>
	    This report's rules are currently calculating total prices
    <% ElseIf strAlertStatus = "T" Then %>
	    This report's rules are currently calculating total prices
    <% ElseIf strAlertStatus = "I" Then %>
	    This report is currently being compared against your rate strategy
    <% ElseIf strAlertStatus = "N" Then %>
	    This report's rules have not yet been processed
    <% ElseIf strAlertStatus = "X" Then %>
	    This report profile does not have any rules assigned
    <% End If %>
	</p>

    
    
    
    &nbsp;</td>
  </tr>
</table>
<p>
<br>


</p>


<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1" width="642">
  <tr>
    <td width="164" bgcolor="#92393A" bordercolor="#92393A" valign="top">
    <font face="Verdana" color="#EAECF9" size="4">&nbsp;Report Legend</font></td>
    <td width="478" bordercolor="#92393A" bgcolor="#92393A" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Profile:</td>
    <td width="478" class="report_detail_light" valign="top"><%=strReportName %></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Locations:</td>
    <td width="478" class="report_detail_light" valign="top"><%=strCityCd %></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Companies: </td>
    <td width="478" class="report_detail_light" valign="top"><%=Left(strVendList, Len(strVendList) - 2)%></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Data Source:</td>
    <td width="478" class="report_detail_light" valign="top"><%=Left(strSourceList, Len(strSourceList) - 2)%></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Dates:</td>
    <td width="478" class="report_detail_light" valign="top"><%=FormatDateTime(varDates(LBound(varDates)+ 1),1) %> to <%=FormatDateTime(varDates(UBound(varDates)),1) %>, 
 	<% Dim strDow
       Dim intDowIndex
       strDow = Split(strDowString, ",")
	   strDowString = ""
       For intDowIndex = LBound(strDow) To UBound(strDow)
       	 Select Case strDow(intDowIndex)
       	 
       	 	Case 1
	        	strDowString = strDowString & "Sun, "
       	 	Case 2
	        	strDowString = strDowString & "Mon, "
       	 	Case 3
	        	strDowString = strDowString & "Tue, "
       	 	Case 4
	        	strDowString = strDowString & "Wed, "
       	 	Case 5
	        	strDowString = strDowString & "Thu, "
       	 	Case 6
	        	strDowString = strDowString & "Fri, "
       	 	Case 7
	        	strDowString = strDowString & "Sat, "
	        	
	     End Select
	        	
	        	
       Next
    %><%=Left(strDowString, Len(strDowString) - 2)%></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Car Types: </td>
    <td width="478" class="report_detail_light" valign="top"><%=Left(strCarList, Len(strCarList) - 2)%>
    </td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Pickup/Drop-off:</td>
    <td width="478" class="report_detail_light">
    <% If strLOR = 1000  Then %>
    	<%="Daily Rates"          %>    
    <% ElseIf strLOR = 1001  Then %>
    	<%="Weekend Daily Rates"         %>    
    <% ElseIf strLOR = 1002  Then %>
    	<%="Weekly Rates"         %>    
    <% Else                   %>
    	<%="LOR " & strLOR         %>
    <% End If                 %>; Pickup <%=FormatDateTime(adoRS.Fields("arv_tm").Value, 3) %>; Drop-off <%=FormatDateTime(adoRS.Fields("rtrn_tm").Value, 3) %></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Comparison company:</td>
    <td width="478" class="report_detail_light" valign="top"> <%=varVendors(1) %></td>
  </tr>
  <tr>
    <td width="164" class="report_detail_dark" valign="top">Rate Displayed:</td>
    <td width="478" class="report_detail_light" valign="top">
    <%=strRateString %>
    </td>
  </tr>
  </table>
<font color="#800000" size="2">
<p><font face="Tahoma">The highlighted vendor is on the far left of the report 
and their rates are always in black. All of the other vendor�s rates will either 
be displayed in red, green, or black. Occasionally you will see other messages 
that appear on reports. The following is a list of each of their meanings:</font></p>
<p><font face="Tahoma" color="#FF0000"><b>Red Rates </b></font>� Rates are 
less than the highlighted vendor's</p>
<p><font face="Tahoma" color="#008000"><b>Green Rates </b></font>� Rates 
are higher than the highlighted vendor's</p>
<p><font face="Tahoma" color="#000000"><b>Black Rates</b></font>
� Rates are the same as the highlighted vendor's</p>
<p><font face="Tahoma"><b>Closed</b> � The car type is not available on a certain 
day for a certain website (i.e. it is sold out).</font></p>
</font>
<p>
&nbsp;</p>


<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="export" width="642">
  <tr>
    <td width="222" bgcolor="#92393A" bordercolor="#92393A" valign="top">
    <font face="Verdana" color="#EAECF9" size="4">&nbsp;Report Utilities</font></td>
    <td width="420" bordercolor="#92393A" bgcolor="#92393A" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td width="222" class="report_detail_dark" valign="top">Download to CSV 
    format:</td>
    <td width="420" class="report_detail_light" valign="top">&nbsp;<a href="car_report_by_type_group_export.asp?reportrequestid=<%=intReportRequestId %>&reportformat=1&security_code=<%=Escape(strSecurityCode) %>">download</a> 
	(please choose an appropriate file name)</td>
  </tr>
  <tr>
    <td width="222" class="report_detail_dark" valign="top">Download to XLS 
    format:</td>
    <td width="420" class="report_detail_light" valign="top">&nbsp;<a href="car_report_by_type_group_export.asp?reportrequestid=<%=intReportRequestId %>&reportformat=0&security_code=<%=Escape(strSecurityCode)%>">download</a> </td>
  </tr>
  <tr>
    <td width="222" class="report_detail_dark" valign="top">Download to XML 
    format:</td>
    <td width="420" class="report_detail_light" valign="top">&nbsp;<a href="car_report_by_type_group_export.asp?reportrequestid=<%=intReportRequestId %>&reportformat=6&security_code=<%=Escape(strSecurityCode)%>">download</a> </td>
  </tr>
  </table>
<font color="#800000" size="2">
<p>Default New User Report - For custom versions please contact your 
Rate-Highway representative at (877) RATE-HWY</p>

</font>
<p><font color="#92393A" style="font-size: 11pt"><b>For more information:</b></font><br>
Click to view the <a href="docs/FAQ_Web.pdf">Web FAQ</a><br>
Click to view the <a href="docs/FAQ_Sabre.pdf">GDS/Sabre FAQ</a></p>
<p>&nbsp;</p>
<p><font style="font-size: 7pt">debug: <%=intDisplayRateType %> / <%=IsNumeric(Request("display_rate_type")) & " us:"%><%=adoRSettings.Fields("us_date").Value %> </font> </p>
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
    Set adoRSettings = Nothing
   
%>
</body>

</html>