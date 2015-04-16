<%@ Language=VBScript %>
<!-- #INCLUDE FILE="include/adovbs.asp" -->
<!-- #INCLUDE FILE="include/DanDate.inc"-->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   On Error Resume Next

   Server.ScriptTimeout = 60
 
	Dim strConn	
	Dim adoCmd	
	Dim adoRS
    Dim lorstr
	Dim adoPrices
	Dim intIndex
	Dim blnExit
	Dim intReportRequestID
	Dim intErrorCount
	Dim intTimeoutCount
	Dim strCarType 
	Dim intResults
	Dim intPrice
    Dim currencyType
	Dim varCarTypes()
	Dim varDataSources()
	Dim varCities()
	Dim varVendors()
    Dim varPodAddress()
	Dim varCurrency()
	Dim varDates()
    Dim varLor()
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
	Dim strIPAddress
	Dim strLOR	
	Dim IsMultiRate
	Dim strRateColumn
	Dim intDisplayRateType 
	Dim blnOnewayReverse 
	Dim intLocale
	Dim strInsurance 
    Dim strProfileCollectionId
    Dim displayRelated
    Dim strSecurityCode
	Dim intDivId

	intDivId = Request("dv")
    strConn = "Provider=SQLOLEDB; Network Library=dbmssocn;Password=iLOVEtab@sco!;User ID=rhWeb;Initial Catalog=shared;Data Source=thor.rate-monitor.com;"
    Set adoCmd = CreateObject("ADODB.Command")
 	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "division_connection_select"
	adoCmd.CommandType = 4
	adoCmd.Parameters.Append adoCmd.CreateParameter("@division_id", 3, 1, 0, intDivId) 
    Set adoRS = Server.CreateObject("ADODB.Recordset")
    adoRS.CursorLocation = adUseClient 
   	adoRS.Open adoCmd, , adOpenStatic, adLockReadOnly
    Session("pro_con") = "Provider=SQLOLEDB; Network Library=dbmssocn;Password=iLOVEtab@sco!;User ID=rhWeb;Initial Catalog=" & adoRS.Fields("database_name").Value & ";Data Source=" & adoRS.Fields("server_name").Value & ".rate-monitor.com;"
    Session("dbserver") = adoRS.Fields("server_name").Value
    intOrgId = adoRS.Fields("org_id").Value
    If adoRS.Fields("us_decimal").Value Then
		intLocale = SetLocale(1033) ' US
	Else
		intLocale = SetLocale(1031) ' Germany
	End If

	strIPAddress = Request.Servervariables("REMOTE_ADDR") 
	intReportRequestId = Request("reportrequestid")
	strSecurityCode =    Request("security_code")
	strCarTypeCd =       Request("car_type_cd")
	strCityCd =          Request("city_cd")
	intDisplayRateType = Request("displayratetype")
	strVendor          = Request("vend_override")
    strProfileCollectionId = Request("profile_collection_id")
    displayRelated      = Request("displayrelated")
    if intOrgId = 148 Then 
        displayRelated = 1 'exception for EHI
    else
        displayRelated = Request("displayrelated")
    end if

	strConn = Session("pro_con")
   	Set adoCmd = CreateObject("ADODB.Command")
  	
	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "currency_select"
	adoCmd.CommandType = 4
	
	Set adoRCurrency = adoCmd.Execute

  	Set adoCmd = CreateObject("ADODB.Command")
  	

  	Set adoRS = CreateObject("ADODB.Recordset")
  	Set adoCmd = CreateObject("ADODB.Command")
	adoRS.CursorLocation = adUseClient
	adoCmd.ActiveConnection = strConn
    if intOrgId = 148 then
        adoCmd.CommandText = "car_shopped_rate_select_rpt4_city"
    else
	    adoCmd.CommandText = "car_shopped_rate_select_rpt4"
    end if
	adoCmd.CommandType = 4
    adoCmd.CommandTimeout=90
	
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
        if intOrgId = 148 Then 
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 6, "ALLLL")	
        Else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 6, Null)	
        End if
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 6, strCityCd)
	End If

    If strProfileCollectionId = "" Then
        adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_collection_id", 3, 1, 0, Null)
    Else
        adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_collection_id", 3, 1, 0, Null)
    End If

    adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_city_cd", 200, 1, 6, Null)
    adoCmd.Parameters.Append adoCmd.CreateParameter("@debug", 3, 1, 0, Null)
    adoCmd.Parameters.Append adoCmd.CreateParameter("@currency_cd_override", 200, 1, 3, Null)

	If Session("user_level") = "0" Then
        adoCmd.Parameters.Append adoCmd.CreateParameter("@security_code", 200, 1, 50, Request.Cookies("password"))
    Else
        adoCmd.Parameters.Append adoCmd.CreateParameter("@security_code", 200, 1, 50, strSecurityCode)
    End If
	adoRS.Open adoCmd, , adOpenStatic, adLockReadOnly, adCmdStoredProc 

	Dim intDateIndex
	Dim intCarTypeIndex
	Dim intDataSourceIndex
	Dim intCityIndex
    Dim intLorIndex

	If adoRS.State = adStateClosed Then
	   Rem If the report request was bogus
	   Server.Transfer "invalid_report.asp"
	End If

	While NOT adoRS.EOF 
		intDateIndex = adoRS.Fields("date_count").Value
		intDataSourceIndex = adoRS.Fields("data_source_count").Value
        intLorIndex =        adoRS.Fields("lor_count").Value
        if intLorIndex = 1 then displayRelated = ""
		intCarTypeIndex =    adoRS.Fields("car_type_count").Value
		intVendorIndex =     adoRS.Fields("vendor_count").Value
		strCarCodeList =     adoRS.Fields("car_type_list").Value
		strDowString =       adoRS.Fields("dow_list").Value	
		strLOR =             adoRS.Fields("lor").Value	
		intCityIndex =       adoRS.Fields("city_count").Value	
		strCityCodeList =    adoRS.Fields("city_cd_list").Value
		If IsNumeric(intDisplayRateType) = False Then
			Rem Allow the user to override the database setting, but if it isn't valid, use the database value
			intDisplayRateType = adoRS.Fields("display_rate_type").Value
		End If
		strAlertStatus =     adoRS.Fields("alert_status").Value
		intRateChanges =     adoRS.Fields("rate_changes").Value	
		tmpDisplayRateType = adoRS.Fields("display_rate_type").Value
		blnOnewayReverse   = adoRS.Fields("oneway_reverse").Value
		strReportName      = adoRS.Fields("rpt_desc").Value
        if Request("largeReportOverride")<>"" Then
            largeReportOverride = Request("largeReportOverride")
        else
            largeReportOverride = adoRS.Fields("large_rpt_override").Value
        end if
		adoRS.MoveNext
	Wend

	strCarCodeListArray = Split(strCarCodeList, ",")
	strCityCodeListArray = Split(strCityCodeList, ",")

	ReDim varCarTypes(intCarTypeIndex)
	ReDim varDataSources(intDataSourceIndex)
	ReDim varDates(intDateIndex)
	ReDim varVendors(intVendorIndex)
	ReDim varVendorCds(intVendorIndex)
    ReDim varPodAddress(intVendorIndex)
	ReDim varCities(intCityIndex)
    ReDim varLor(intLorIndex)
	
	intDateIndex = 0
	intCarTypeIndex = 0
	intDataSourceIndex = 0
	intVendorIndex = 0
	intCityIndex = 0
	
	Set adoRS = adoRS.NextRecordset
    Dim arv_tm 
    arv_tm = adoRS.Fields("arv_tm").Value
    Dim rtrn_tm
    rtrn_tm = adoRS.Fields("rtrn_tm").Value

	While adoRS.EOF = False
 	   If intDateIndex < UBound(varDates) Then
		   If varDates(intDateIndex) <> adoRS.Fields("arv_dt").Value Then
			   intDateIndex= intDateIndex+ 1
	    	   If intDateIndex <= UBound(varDates) Then
					varDates(intDateIndex) = adoRS.Fields("arv_dt").Value
				End If
			End if
		End If

 	   If intLorIndex < UBound(varLor) Then
		   If varLor(intLorIndex) <> adoRS.Fields("lor").Value Then
			   intLorIndex = intLorIndex + 1
	    	   If intLorIndex <= UBound(varLor) Then
					varLor(intLorIndex) = adoRS.Fields("lor").Value
				End If
 			End if
		End If

		If intCarTypeIndex < UBound(varCarTypes) Then
			If varCarTypes(intCarTypeIndex ) <> adoRS.Fields("shop_car_type_cd").Value Then
				intCarTypeIndex = intCarTypeIndex + 1
				If intCarTypeIndex <= UBound(varCarTypes) Then
					varCarTypes(intCarTypeIndex ) = adoRS.Fields("shop_car_type_cd").Value 
					strCarList = strCarList & adoRS.Fields("shop_car_type_cd").Value & ", " 
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
 if intOrgId = 148 then varPodAddress(intVendorIndex) = adoRS.Fields("pod_street_addr").Value
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

 	Rem Allow the user to override this
	If IsNumeric(Request("display_rate_type")) Then
		If Request("display_rate_type") > 0 And Request("display_rate_type") < 5 Then
			intDisplayRateType = Request("display_rate_type")
		Else
			Rem otherwise leave it as is from the database
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


	If err.number <> 0 Then
	'   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	'   response.write "<b>VBScript Errors Occured!<br>"
	'   response.write parm_msg & "</b><br>"
	 '  response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	'   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	'   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	'   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	 '  response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If

    adoRS.MoveFirst
	%>
<html>

<head>
<link rel="SHORTCUT ICON" href="http://www.rate-highway.com/favicon.ico">
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate-Monitor.com | Rate Report by Type</title>
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<script type="text/javascript" src="js/jquery-1.10.2.min.js"></script>
<script language='Javascript' type="text/javascript" >

    var userLang = (navigator.language) ? navigator.language : navigator.userLanguage;
    //alert ("The language is: " + userLang);

    function centerPopUp(url, name, width, height, scrollbars) {

        if (scrollbars == null) scrollbars = "0"

        str = "";
        str += "resizable=1,";
        str += "scrollbars=" + scrollbars + ",";
        str += "width=" + width + ",";
        str += "height=" + height + ",";

        if (window.screen) {
            var ah = screen.availHeight - 30;
            var aw = screen.availWidth - 10;

            var xc = (aw - width) / 2;
            var yc = (ah - height) / 2;

            str += ",left=" + xc + ",screenX=" + xc;
            str += ",top=" + yc + ",screenY=" + yc;
        }
        window.open(url, name, str);
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

<body leftmargin="0" topmargin="0" onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img alt="" src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">
    <!-- #INCLUDE FILE="include/page_header_buttons.htm" -->
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/med_bar_tile.gif">
    <img src="images/med_bar.gif" width="12" height="8" alt=""></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/user_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/user_left.gif" width="580" height="31" alt=""></td>
        <td background="images/user_tile.gif">
        <table width="100" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td valign="bottom">
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
                <div align="right">
                  <font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
                  User: <%=Request.Cookies("user_name")%></font></div>
                </td>
              </tr>
              <tr>
                <td><img src="images/separator.gif" width="183" height="6" alt=""></td>
              </tr>
            </table>
            </td>
            <td><img src="images/user_tile.gif" width="7" height="31" alt=""></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
	
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#92393A" width="100%" id="AutoNumber2">
      <tr>
        <td width="100%" bgcolor="#92393A">&nbsp;<font color="#FFFFFF" size="4" face="Verdana">Rate Detail 
		- <%=strReportName %> <%=lorstr%></font></td>
      </tr>
</table>

<table cellSpacing="0" cellPadding="8" width="100%" border="0">
  <tr>
    <td bgColor="#CFD7DB" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif; width: 874px;">
		<form method="GET" name="rate_detail">
        <input type="hidden" name="largeReportOverride" value="<%=largeReportOverride %>" />
        <input type="hidden" name="dv" value="<%=intUserId%>" />
			<% if intOrgId = 148 then %>
            <input type="hidden" name="city_cd" value="ALLLL" />
            <%else %>
                  <select size="1" name="city_cd" title="City to display">
				  <% For intIndex = LBound(strCityCodeListArray) To UBound(strCityCodeListArray)	%>
				  	<% If strCityCd = strCityCodeListArray(intIndex) Then             %>
				   				   <option value="<%=trim(strCityCodeListArray(intIndex))%>" selected><%=trim(strCityCodeListArray(intIndex))%></option>
				  	<% Else             %>
				   				   <option value="<%=trim(strCityCodeListArray(intIndex))%>"><%=trim(strCityCodeListArray(intIndex))%></option>
				  	<% End If           %>

				  <% Next %>  				 
 
                    <% if largeReportOverride = false or largeReportOverride = "False" Then %>
			      <% If strCityCd = "ALLLL" Then %>
				   <option selected value="ALLLL" >ALL </option>
                  <% Else %>
				   <option value="ALLLL" >ALL </option>
                  <% End If %>	
                  <% End If %>
            <%end if %>
		           </select>
                  <select size="1" name="car_type_cd" title="Car type to display">
				  <% For intIndex = LBound(strCarCodeListArray) To UBound(strCarCodeListArray)	%>
				  <% If strCarTypeCd = strCarCodeListArray(intIndex) Then             %>
				   				   <option value="<%=trim(strCarCodeListArray(intIndex))%>" selected><%=trim(strCarCodeListArray(intIndex))%></option>
				  <% Else             %>
				   				   <option value="<%=trim(strCarCodeListArray(intIndex))%>"><%=trim(strCarCodeListArray(intIndex))%></option>
				  <% End If           %>
 
				  <% Next %>                  
                  
                  <% if largeReportOverride = false or largeReportOverride = "False" Then %>
                  <% If strCarTypeCd = "ALLL" Then %>
				   <option selected value="ALLL" >ALL </option>
                  <% Else %>
				   <option value="ALLL" >ALL </option>
                  <% End If %>
                  <% End If %>	
                  </select> 
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
				  </select>
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
				  </select>
				  <% End If %> 
				  <select size="1" name="displayratetype" style="width:225" title="Rate type to display">
                    <option value="1"<%if intDisplayRateType = 1 then Response.Write " selected"%>>Base rate amt</option>
                    <!--<option value="2"<%if intDisplayRateType = 2 then Response.Write " selected"%>>Total rate amount w/o fees</option>-->
                    <option value="3"<%if intDisplayRateType = 3 then Response.Write " selected"%>>Total price w/ fees</option>
                    <option value="4"<%if intDisplayRateType = 4 then Response.Write " selected"%>>Rate amt &amp; Total price</option>
                    <!--<option value="5"<%if intDisplayRateType = 5 then Response.Write " selected"%>>Rate amt/Total price/Drop Chg</option>
                    <option value="6"<%if intDisplayRateType = 6 then Response.Write " selected"%>>Rate amt/Total price/Extra day</option>
                    <option value="7"<%if intDisplayRateType = 7 then Response.Write " selected"%>>Rate amt/Total price/Limit/Extra</option>
                    <option value="8"<%if intDisplayRateType = 8 then Response.Write " selected"%>>Total price/Ins. Included</option>
                    <option value="9"<%if intDisplayRateType = 9 then Response.Write " selected"%>>Total price/Ins. Not Included</option>-->
                     </select>
                  <input type="submit" value="Display" name="display" class="rh_button" id="display" style="border:3px double #2A2F34; font-family: Vendana, Arial, Helvetica, sans-serif; font-size: 10pt; color:#FFFFFF; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1; background-color:#384F5B; font-weight:bold" >
                  <input type="hidden" name="reportrequestid" value="<%=Request("reportrequestid") %>">
                  <input type="hidden" name="security_code" value="<%=strSecurityCode %>">
                  <%if intLorIndex > 1 Then %>
                    <input type="submit" value="Display Related LORs" name="displayrelated" class="rh_button" id="Submit2" style="border:3px double #2A2F34; font-family: Vendana, Arial, Helvetica, sans-serif; font-size: 10pt; color:#FFFFFF; padding-left:4; padding-right:4; padding-top:1; padding-bottom:1; background-color:#384F5B; font-weight:bold" >
                  <%End if %>
				  <br> </form> 
         <%if largeReportOverride=true or largeReportOverride="True" then Response.Write "Your report is too large to display multiple locations or car types<br>please select the location and car type you wish to view." %>
   </td>
    <td vAlign="top" bgColor="#CFD7DB" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif" >
    &nbsp;&nbsp; 
    
    &nbsp;</td>
  </tr>
</table>
    
<div id="multilor"></div>

<script>
//    $('#multilor').load('car_report_by_type.aspx?largeReportOverride=<%=Request("largeReportOverride")%>&city_cd=<%=Request("city_cd")%>&car_type_cd=<%=Request("car_type_cd")%>&vend_override=<%=Request("vend_override")%>&currency_override=<%=Request("currency_override")%>&displayratetype=<%=Request("displayratetype")%>4&display=<%=Request("display")%>&ReportRequestID=<%=Request("ReportRequestID") %>&dv=<%=Request("dv")%>&security_code=<%=Request("security_code")%>&output_style=1');
    $('#multilor').load('car_report_by_type.aspx?ReportRequestID=<%=Request("ReportRequestID") %>&dv=<%=Request("dv")%>&security_code=<%=Request("security_code")%>&output_style=1');
</script>

	
	<%

	While NOT adoRS.EOF 
		If (strCarType <> adoRS.Fields("shop_car_type_cd").Value) Or _
		   ((strCityCd <> adoRS.Fields("city_cd").Value & " to " & adoRS.Fields("rtrn_city_cd").Value & "") And (strCityCd <> adoRS.Fields("city_cd").Value)) Then
			
			strCarType = adoRS.Fields("shop_car_type_cd").Value
			If (IsNull(adoRS.Fields("rtrn_city_cd").Value)) Or (adoRS.Fields("city_cd").Value = adoRS.Fields("rtrn_city_cd").Value & "") Then
				strCityCd = adoRS.Fields("city_cd").Value 
			Else
				strCityCd = adoRS.Fields("city_cd").Value & " to " & adoRS.Fields("rtrn_city_cd").Value
			End If
            counter = 0
			%>
			<table cellSpacing='0' cellPadding='8' width='100%' border='0'>
			  <tr>
			    <td vAlign="top" bgColor="#CFD7DB" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">    	
                  <p><font size="+1"><b><%=strCityCd %></b></font>, <font size="+1"><b><%=strCarType %></b><%=strRateTitle %>       
					</font><table cellSpacing="0" cellPadding="2" border="1">

            	<%
                blnDarkRow = True
		      	
				For intIndex = 1 To intDateIndex 
					blnMoved = False
                if displayRelated <> "" or counter = 0 Then

                %>
		      <tr>
		      <%
              if displayRelated = "" then %>
  		      <th noWrap align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif">&nbsp;</th>
              <% else
                    If adoRSettings.Fields("us_date").Value Then %>
		      <th noWrap align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=WeekDayName(Weekday(adoRS.Fields("arv_dt").Value ), True) & " - " & FormatDateTime(adoRS.Fields("arv_dt").Value , 2) %>&nbsp;</th>
		      <% Else %>
		      <th noWrap align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=DanDate(adoRS.Fields("arv_dt").Value , "%d/%m/%y") %>&nbsp;</th>
		      <% End If 
                end if%>
		      	<%
		      	For intVendorIndex = LBound(varVendors) + 1  To UBound(varVendors) 
		      		blnMoved = False
                    
		      	%>
		        <th align="middle" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif" title="<%=varPodAddress(intVendorIndex) %>"><%=varVendors(intVendorIndex) %>&nbsp;</th>
		      	<%
		      	Next
		      	%>

        
		      </tr>
		      <%
                 End If
                        counter = 1
               For i = 1 to intLorIndex
			   
                        blnDarkRow = Not blnDarkRow
  					
					If adoRS.EOF = True Then
						exit for
					end if
					
					curRate = adoRS.Fields(strRateColumn).Value
					curTotal = adoRS.Fields("est_rental_chrg_amt").Value				
if strLOR = adoRS.Fields("lor") or displayRelated <> "" Then
		      %>
		      <tr>
		      <% 
 
				
				

                if displayRelated = "" then
 		            If adoRSettings.Fields("us_date").Value Then %>
		      <th noWrap align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=WeekDayName(Weekday(adoRS.Fields("arv_dt").Value ), True) & " - " & FormatDateTime(adoRS.Fields("arv_dt").Value , 2) %>&nbsp;</th>
		            <% Else %>
		      <th noWrap align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=DanDate(adoRS.Fields("arv_dt").Value , "%d/%m/%y") %>&nbsp;</th>
		            <% End If 
               else
		      %>
		      <th noWrap align="right" bgColor="#B2BEC4" style="font-size: 0.8em; color: navy; font-family: Verdana, Arial, sans-serif"><%=adoRS.Fields("lor") %>LOR&nbsp;</th>
	      
			  
		      <% 
                end if
 				For intVendorIndex = LBound(varVendors) To UBound(varVendors) '- 1	
                        Select Case adoRS.Fields("currency_cd").Value
                            Case "GBP"
                                curSym = "&pound;"
                            Case "AED"
                                curSym = "AED;"
				      		Case Else 
                                curSym = ""
			      		End Select 

						If blnDarkRow Then
							strBgColor = "#B2BEC4"
						Else
							strBgColor = "#CFD7DB"
						End If					
					
					If adoRS.EOF = True Then
						Exit For
					End If   
					
					
					If (intVendorIndex = UBound(varVendors)) Then
		
	      
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
		    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif">Blocked</td>
						  <%


		      			ElseIf curRate > adoRS.Fields(strRateColumn).Value Then
		      			
					      %>
					      	<% Select Case intDisplayRateType %>
					      		<% Case 4 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 5 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & curSym & FormatNumber("0" & adoRS.Fields("extra_values").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 6 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & curSym & FormatNumber("0" & adoRS.Fields("extra_da_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 7 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & adoRS.Fields("rent_mi_alwnc").Value & "/" & adoRS.Fields("mi_chrg_amt").Value %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 8 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %>
				    			    <%=strInsurance %></td>

				      			<% Case Else %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: red; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %></td>
			      			<% End Select %>
						  <%
						  
						ElseIf curRate < adoRS.Fields(strRateColumn).Value Then
						
					      %>
					      	<% Select Case intDisplayRateType %>
					      		<% Case 4 %>
							        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 5 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & curSym & FormatNumber("0" & adoRS.Fields("extra_values").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 6 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & curSym & FormatNumber("0" & adoRS.Fields("extra_da_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 7 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & adoRS.Fields("rent_mi_alwnc").Value & "/" & adoRS.Fields("mi_chrg_amt").Value %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 8 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %>
				    			    <%=strInsurance %></td>

					      		<% Case Else %>
							        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: green; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %></td>
				      		<% End Select %>
						  <%
						  
						Else
					      %>
					      	<% Select Case intDisplayRateType %>
					      		<% Case 4 %>
							        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 5 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & curSym & FormatNumber("0" & adoRS.Fields("extra_values").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 6 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & curSym & FormatNumber("0" & adoRS.Fields("extra_da_chrg_amt").Value, 2) %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 7 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields("rt_amt").Value, 2) & "/" & curSym & FormatNumber(adoRS.Fields("est_rental_chrg_amt").Value, 2) & "/" & adoRS.Fields("rent_mi_alwnc").Value & "/" & adoRS.Fields("mi_chrg_amt").Value %>
					      			<sub><%=trim(adoRS.Fields("rt_type_cd").Value & " " & adoRS.Fields("rate_cd").Value) %></sub></td>
					      		<% Case 8 %>
				    			    <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %>
				    			    <%=strInsurance %></td>

					      		<% Case Else %>
							        <td noWrap align="right" bgColor="<%=strBgColor  %>" style="font-size: 0.8em; color: black; font-family: Verdana, Arial, sans-serif"><%=strMilage %><%= curSym & FormatNumber(adoRS.Fields(strRateColumn).Value, 2) %></td>
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
 Else
    ///JUMP AHEAD TO THE NEXT SET OF DATA
    For ii = 1 to intVendorIndex - 1
        adoRS.MoveNext
    Next
End If
                Next
		      	Next 
		      %>
		      
      		</table>
      		</p></td>
      	</table>
			<%
			
		End If
		
		'adoRS.MoveNext
	Wend


rem	If err.number <> 0 Then
rem	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
rem	   response.write "<b>VBScript Errors Occured!<br>"
rem	   response.write parm_msg & "</b><br>"
rem	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
rem	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
rem	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
rem	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
rem	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
rem	End If

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

<%if intOrgId = 148 then %>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="Table1" width="642">
  <tr>
    <td width="164" bgcolor="#92393A" bordercolor="#92393A" valign="top">
    <font face="Verdana" color="#EAECF9" size="4">&nbsp;Location Key</font></td>
    <td width="478" bordercolor="#92393A" bgcolor="#92393A" valign="top">&nbsp;</td>
  </tr>
<%for i=1 to intVendorIndex-1 %>
  <tr>
    <td width="164" class="report_detail_dark" valign="top"><%=varVendors(i) %>:</td>
    <td width="478" class="report_detail_light" valign="top"><%=varPodAddress(i) %></td>
  </tr>
<%Next %>
</table>
<%end if %>

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
    <% End If                 %>; Pickup <%=FormatDateTime(arv_tm, 3) %>; Drop-off <%=FormatDateTime(rtrn_tm, 3) %></td>
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
and their rates are always in black. All of the other vendor’s rates will either 
be displayed in red, green, or black. Occasionally you will see other messages 
that appear on reports. The following is a list of each of their meanings:</font></p>
<p><font face="Tahoma" color="#FF0000"><b>Red Rates </b></font>– Rates are 
less than the highlighted vendor's</p>
<p><font face="Tahoma" color="#008000"><b>Green Rates </b></font>– Rates 
are higher than the highlighted vendor's</p>
<p><font face="Tahoma" color="#000000"><b>Black Rates</b></font>
– Rates are the same as the highlighted vendor's</p>
<p><font face="Tahoma"><b>Closed</b> – The car type is not available on a certain 
day for a certain website (i.e. it is sold out). </font></p>
</font>
<p>
&nbsp;</p>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="export" width="642">
  <tr>
    <td width="222" bgcolor="#92393A" bordercolor="#92393A" valign="top">
    <font face="Verdana" color="#EAECF9" size="4">&nbsp;Report Utilities</font></td>
    <td width="420" bordercolor="#92393A" bgcolor="#92393A" valign="top">&nbsp;</td>
  </tr>
 <!-- <tr>
    <td width="222" class="report_detail_dark" valign="top">Download to CSV 
    format:</td>
    <td width="420" class="report_detail_light" valign="top">&nbsp;<a href="car_report_by_type_export.asp?reportrequestid=<%=intReportRequestId %>&reportformat=1&security_code=<%=Escape(strSecurityCode) %>">download</a> 
	(please choose an appropriate file name)</td>
  </tr>-->
  <tr>
    <td width="222" class="report_detail_dark" valign="top">Download to Excel 
    format:</td>
    <td width="420" class="report_detail_light" valign="top">&nbsp;<a href="car_report_by_type_export<%=Request.QueryString("large") %>.asp?reportrequestid=<%=intReportRequestId %>&reportformat=7&security_code=<%=Escape(strSecurityCode)%>">download</a> </td>
  </tr>
 <!-- <tr>
    <td width="222" class="report_detail_dark" valign="top">Download to XML 
    format:</td>
    <td width="420" class="report_detail_light" valign="top">&nbsp;<a href="car_report_by_type_export.asp?reportrequestid=<%=intReportRequestId %>&reportformat=6&security_code=<%=Escape(strSecurityCode)%>">download</a> </td>
  </tr>-->
  </table>

<%
rem	If err.number <> 0 Then
rem	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
rem	   response.write "<b>VBScript Errors Occured!<br>"
rem	   response.write parm_msg & "</b><br>"
rem	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
rem	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
rem	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
rem	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
rem	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
rem	End If

    Set adoRS = Nothing 
    Set adoCmd = Nothing 
    Set adoRSettings = Nothing
   
%>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>